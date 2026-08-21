require('dotenv').config();
const crypto = require('crypto');
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const cookieParser = require('cookie-parser');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(cors({
  origin: true,
  credentials: true,
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(cookieParser());
const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET) {
  console.error('❌ JWT_SECRET не задан в переменных окружения');
  process.exit(1);
}

let credentials;
try {
  const keyFile = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  credentials = JSON.parse(fs.readFileSync(keyFile));
  console.log('✅ Сервисный аккаунт:', credentials.client_email);
} catch (err) {
  console.error('❌ Ошибка загрузки credentials.json:', err.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

async function getSheetData(range) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
    });
    return res.data.values || [];
  } catch (err) {
    console.error(`Ошибка чтения ${range}:`, err.message);
    return [];
  }
}

async function appendRow(range, row) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      resource: { values: [row] },
    });
    return true;
  } catch (err) {
    console.error(`Ошибка добавления строки в ${range}:`, err.message);
    return false;
  }
}

async function updateRow(range, rowNumber, values) {
  console.log(`🔄 updateRow вызвана: лист="${range}", строка=${rowNumber}, значения:`, values);
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${range}!A${rowNumber}:${String.fromCharCode(64 + values.length)}${rowNumber}`,
      valueInputOption: 'RAW',
      resource: { values: [values] },
    });
    console.log(`✅ updateRow успешно обновила строку ${rowNumber} в листе ${range}`);
    return true;
  } catch (err) {
    console.error(`❌ Ошибка обновления строки ${rowNumber}:`, err.message);
    return false;
  }
}

async function deleteRow(sheetName, rowNumber) {
  console.log(`🔄 deleteRow: лист="${sheetName}", строка=${rowNumber}`);
  const sheetId = await getSheetId(sheetName);
  if (!sheetId) {
    console.error(`❌ Не найден sheetId для "${sheetName}"`);
    return false;
  }
  console.log(`✅ sheetId = ${sheetId}`);
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowNumber - 1,
              endIndex: rowNumber,
            }
          }
        }]
      }
    });
    console.log(`✅ Строка ${rowNumber} удалена`);
    return true;
  } catch (err) {
    console.error(`❌ Ошибка удаления строки ${rowNumber}:`, err.message);
    return false;
  }
}

async function getSheetId(sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetName);
  return sheet ? sheet.properties.sheetId : null;
}

function authenticateToken(req, res, next) {
  // Сначала проверим заголовок Authorization
  const authHeader = req.headers.authorization;
  let token = authHeader && authHeader.split(' ')[1];
  // Если нет в заголовке, проверим куку (для совместимости)
  if (!token) token = req.cookies?.access_token;
  if (!token) return res.status(401).json({ error: 'Не авторизован' });
  try {
    const user = jwt.verify(token, JWT_SECRET);
    req.user = user;
    next();
  } catch {
    res.status(403).json({ error: 'Недействительный токен' });
  }
}

app.get('/ping', (req, res) => {
  res.json({ message: 'Сервер работает' });
});

// ---------- ЗАКАЗЫ ----------
app.get('/api/orders', async (req, res) => {
  const rows = await getSheetData('ЗАКАЗЫ!A2:I');
  const orders = rows.map(row => ({
    id: row[0],
    created: row[1],
    clientId: row[2],
    price: parsePrice(row[3]),
    status: row[4],
    details: row[5],
    delivery: row[6],
    executionDate: row[7],
  }));
  res.json(orders);
});

app.post('/api/orders', authenticateToken, async (req, res) => {
  const { clientName, clientPhone, clientAddress, price, status, details, delivery, executionDate } = req.body;

  const clientId = await findOrCreateClient(clientName, clientPhone, clientAddress);

  const rows = await getSheetData('ЗАКАЗЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const now = new Date().toISOString();

  console.log('🔍 clientId из findOrCreateClient:', clientId, typeof clientId);
  const success = await appendRow('ЗАКАЗЫ!A:I', [
    newId, now, clientId !== null ? clientId : '', price, status, details, delivery, executionDate, req.user.userId
  ]);
  if (!success) return res.status(500).json({ error: 'Ошибка создания заказа' });

  if (status === 'оплачен' || status === 'завершен') {
    await createTransactionForOrder(newId, price, clientId);
  }

  const products = await getOrderProductsWithNames(newId);
  await saveOrderSnapshot(newId, req.user.userId, req.user.name, products);

  res.json({ success: true, orderId: newId });
});

app.put('/api/orders/:id', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.id);
  const updates = req.body;
  console.log('📥 Получен запрос на обновление заказа:', req.body);

  const rows = await getSheetData('ЗАКАЗЫ!A2:I');
  let rowIndex = -1;
  let oldRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === orderId) {
      rowIndex = i + 2;
      oldRow = rows[i];
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Заказ не найден' });

  const oldStatus = oldRow[4];
  const newStatus = updates.status;
  const newPrice = updates.price !== undefined ? parseFloat(updates.price) : parseFloat(oldRow[3]);

  console.log('🔍 Полученные данные клиента из запроса:', {
    clientName: updates.clientName,
    clientPhone: updates.clientPhone,
    clientAddress: updates.clientAddress
  });
  let finalClientId = oldRow[2] && oldRow[2] !== '' ? parseInt(oldRow[2]) : null;

  const hasName = updates.clientName !== undefined;
  const hasPhone = updates.clientPhone !== undefined;
  const hasAddress = updates.clientAddress !== undefined;

  if (hasName || hasPhone || hasAddress) {
    let newName = hasName ? (updates.clientName ? updates.clientName.trim() : null) : null;
    let newPhone = hasPhone ? updates.clientPhone.trim() : null;
    let newAddress = hasAddress ? updates.clientAddress.trim() : null;

    if (newName === '' || newName === 'Аноним') newName = null;
    if (newPhone === '' || newPhone === 'не указан') newPhone = null;
    if (newAddress === '') newAddress = null;

    if (newName === null && newPhone === null && newAddress === null) {
      finalClientId = null;
      console.log('Все поля клиента пусты – отвязываем заказ');
    } else {
      console.log('Ищем/создаём клиента с данными:', { newName, newPhone, newAddress });
      finalClientId = await findOrCreateClient(newName, newPhone, newAddress);
      console.log('Результат findOrCreateClient: clientId =', finalClientId);
    }
  }

  const newRow = [
    Number(oldRow[0]),
    oldRow[1],
    finalClientId,
    newPrice,
    newStatus !== undefined ? newStatus : oldRow[4],
    updates.details !== undefined ? updates.details : oldRow[5],
    updates.delivery !== undefined ? updates.delivery : oldRow[6],
    updates.executionDate !== undefined ? updates.executionDate : oldRow[7],
    req.user.userId,
  ];

  const success = await updateRow('ЗАКАЗЫ', rowIndex, newRow);
  if (!success) return res.status(500).json({ error: 'Ошибка обновления заказа' });

  const shouldCreate = (oldStatus !== 'оплачен' && oldStatus !== 'завершен') &&
                       (newStatus === 'оплачен' || newStatus === 'завершен');
  const shouldDelete =
    (oldStatus === 'оплачен' && newStatus !== 'оплачен' && newStatus !== 'завершен') ||
    (oldStatus === 'завершен' && (newStatus === 'в работе' || newStatus === 'отменён'));

  if (shouldCreate) {
    await createTransactionForOrder(orderId, newPrice, finalClientId);
  } else if (shouldDelete) {
    await deleteTransactionByOrderId(orderId);
  }

  const products = await getOrderProductsWithNames(orderId);
  await saveOrderSnapshot(orderId, req.user.userId, req.user.name, products);
  console.log('✅ Заказ обновлён, новый clientId =', finalClientId);
  res.json({ success: true });
});

app.delete('/api/orders/:id', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.id);

  const ordersData = await getSheetData('ЗАКАЗЫ!A2:I');
  let rowIndex = -1;
  let orderRow = null;
  for (let i = 0; i < ordersData.length; i++) {
    if (parseInt(ordersData[i][0]) === orderId) {
      rowIndex = i + 2;
      orderRow = ordersData[i];
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Заказ не найден' });

  const orderStatus = orderRow[4];

  if (orderStatus === 'оплачен') {
    await deleteTransactionByOrderId(orderId);
  }

  const products = await getOrderProductsWithNames(orderId);
  await saveOrderSnapshot(orderId, req.user.userId, req.user.name, products);

  const orderProductsRows = await getSheetData('ЗАКАЗЫ_ТОВАРЫ!A2:D');
  for (let i = 0; i < orderProductsRows.length; i++) {
    if (parseInt(orderProductsRows[i][0]) === orderId) {
      await deleteRow('ЗАКАЗЫ_ТОВАРЫ', i + 2);
    }
  }

  const success = await deleteRow('ЗАКАЗЫ', rowIndex);
  if (!success) return res.status(500).json({ error: 'Ошибка удаления заказа' });

  res.json({ success: true });
});

// ---------- КЛИЕНТЫ ----------
app.get('/api/clients', async (req, res) => {
  const rows = await getSheetData('КЛИЕНТЫ!A2:E');
  const clients = rows.map(row => ({
    id: row[0],
    name: row[1],
    phone: row[2],
    address: row[3],
    notes: row[4],
  }));
  res.json(clients);
});

app.post('/api/clients', authenticateToken, async (req, res) => {
  const { name, phone, address, notes } = req.body;
  const rows = await getSheetData('КЛИЕНТЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const success = await appendRow('КЛИЕНТЫ!A:E', [newId, name, phone || '', address || '', notes || '']);
  if (success) res.json({ success: true, id: newId });
  else res.status(500).json({ error: 'Ошибка создания клиента' });
});

app.put('/api/clients/:id', authenticateToken, async (req, res) => {
  const clientId = parseInt(req.params.id);
  const { name, phone, address, notes } = req.body;

  const rows = await getSheetData('КЛИЕНТЫ!A2:E');
  let rowIndex = -1;
  let oldRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === clientId) {
      rowIndex = i + 2;
      oldRow = rows[i];
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Клиент не найден' });

  const newRow = [
    clientId,
    name !== undefined ? name : oldRow[1],
    phone !== undefined ? phone : oldRow[2],
    address !== undefined ? address : oldRow[3],
    notes !== undefined ? notes : oldRow[4],
  ];

  const success = await updateRow('КЛИЕНТЫ', rowIndex, newRow);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка обновления клиента' });
});

// ---------- СКЛАД ----------
app.get('/api/stock', async (req, res) => {
  const rows = await getSheetData('СКЛАД!A2:D');
  const stock = rows.map(row => ({ id: row[0], name: row[1], stock: row[3] }));
  res.json(stock);
});

app.post('/api/stock', authenticateToken, async (req, res) => {
  res.status(501).json({ error: 'Not implemented' });
});

// ---------- РЕЦЕПТЫ ----------
app.get('/api/recipes', async (req, res) => {
  const rows = await getSheetData('РЕЦЕПТЫ!A2:D');
  const recipes = rows.map(row => ({ id: row[0], name: row[1], yield: row[2], cost: row[3] }));
  res.json(recipes);
});

// ---------- ФИНАНСЫ ----------
app.get('/api/finance', async (req, res) => {
  const { startDate, endDate } = req.query;
  const rows = await getSheetData('ФИНАНСЫ!A2:G');
  let transactions = rows.map(row => {
    const amountStr = row[4] || '0';
    // Заменяем запятую на точку и парсим как число
    const amount = parseFloat(amountStr.replace(',', '.'));
    return {
      id: row[0],
      date: row[1],
      type: row[2],
      category: row[3],
      amount: isNaN(amount) ? 0 : amount,
      comment: row[5],
      orderId: row[6] ? parseInt(row[6]) : null,
    };
  });

  const totalIncome = transactions.filter(t => t.type === 'Доход').reduce((s, t) => s + t.amount, 0);
  const totalExpense = transactions.filter(t => t.type === 'Расход').reduce((s, t) => s + t.amount, 0);

  res.json({
    transactions,
    totalIncome,
    totalExpense,
    profit: totalIncome - totalExpense,
  });
});

app.post('/api/transactions', authenticateToken, async (req, res) => {
  const { type, category, amount, comment, date } = req.body;

  // Преобразуем сумму в число (заменяем запятую на точку)
  let numericAmount;
  if (typeof amount === 'string') {
    numericAmount = parseFloat(amount.replace(',', '.'));
  } else {
    numericAmount = parseFloat(amount);
  }

  if (isNaN(numericAmount) || numericAmount <= 0) {
    return res.status(400).json({ error: 'Сумма должна быть положительным числом (например, 10.50)' });
  }

  const rows = await getSheetData('ФИНАНСЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const now = date ? new Date(date).toISOString() : new Date().toISOString();

  const success = await appendRow('ФИНАНСЫ!A:F', [
    newId,
    now,
    type,
    category,
    numericAmount,  // ← передаём число
    comment || ''
  ]);

  if (success) {
    res.json({ success: true });
  } else {
    res.status(500).json({ error: 'Ошибка добавления транзакции' });
  }
});

// ---------- АВТОРИЗАЦИЯ ----------
app.post('/api/register', async (req, res) => {
  const { name, email, password } = req.body;
  if (!name || !email || !password) {
    return res.status(400).json({ error: 'Заполните все поля' });
  }
  const users = await getSheetData('ПОЛЬЗОВАТЕЛИ!A2:G');
  if (users.some(u => u[2] === email)) {
    return res.status(400).json({ error: 'Пользователь уже существует' });
  }
  const saltRounds = 10;
  const passwordHash = await bcrypt.hash(password, saltRounds);
  let lastId = 0;
  users.forEach(u => { const id = parseInt(u[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const success = await appendRow('ПОЛЬЗОВАТЕЛИ!A:G', [newId, name, email, 'user', passwordHash, '', '']);
  if (success) res.json({ success: true, userId: newId });
  else res.status(500).json({ error: 'Ошибка регистрации' });
});

app.post('/api/login', async (req, res) => {
  const { email, password } = req.body;
  const users = await getSheetData('ПОЛЬЗОВАТЕЛИ!A2:G');
  const user = users.find(u => u[2] === email);
  if (!user) return res.status(401).json({ error: 'Неверные учётные данные' });
  const valid = await bcrypt.compare(password, user[4]);
  if (!valid) return res.status(401).json({ error: 'Неверные учётные данные' });
  const token = jwt.sign(
    { userId: user[0], name: user[1], email, role: user[3] },
    JWT_SECRET,
    { expiresIn: '30d' }
  );
  res.json({ success: true, name: user[1], role: user[3], email, token });
});

app.get('/api/me', authenticateToken, (req, res) => {
  res.json(req.user);
});

app.post('/api/logout', (req, res) => {
  res.clearCookie('access_token');
  res.json({ success: true });
});

// ---------- РАСЧЁТ РЕЦЕПТА ----------
app.post('/api/recipes/calculate', async (req, res) => {
  const { recipeId, desiredWeight } = req.body;
  if (!recipeId || !desiredWeight) {
    return res.status(400).json({ error: 'recipeId и desiredWeight обязательны' });
  }
  try {
    const recipes = await getSheetData('РЕЦЕПТЫ!A2:D');
    const recipe = recipes.find(r => parseInt(r[0]) === recipeId);
    if (!recipe) return res.status(404).json({ error: 'Рецепт не найден' });
    const recipeYield = parseFloat(recipe[2]);
    if (!recipeYield) return res.status(400).json({ error: 'В рецепте не указан выход' });

    const coefficient = desiredWeight / recipeYield;

    const composition = await getSheetData('СОСТАВ_РЕЦЕПТА!A:C');
    const recipeIngredients = composition.filter(row => parseInt(row[0]) === recipeId);
    if (recipeIngredients.length === 0) {
      return res.status(404).json({ error: 'В рецепте нет ингредиентов' });
    }

    const ingredients = await getSheetData('ИНГРЕДИЕНТЫ!A:E');
    const ingMap = new Map();
    ingredients.forEach(ing => ingMap.set(parseInt(ing[0]), { name: ing[1], unit: ing[3], price: parseFloat(ing[4]) || 0 }));

    const resultIngredients = [];
    for (const comp of recipeIngredients) {
      const ingId = parseInt(comp[1]);
      const amount = parseFloat(comp[2]);
      const needed = amount * coefficient;
      const ing = ingMap.get(ingId);
      if (ing) {
        resultIngredients.push({ name: ing.name, neededGrams: needed });
      } else {
        resultIngredients.push({ name: `Ингредиент ID ${ingId}`, neededGrams: needed });
      }
    }

    res.json({ coefficient, ingredients: resultIngredients });
  } catch (err) {
    console.error('Ошибка при расчёте рецепта:', err);
    res.status(500).json({ error: 'Внутренняя ошибка сервера' });
  }
});

app.post('/api/recipes', authenticateToken, async (req, res) => {
  const { name, yield: recipeYield, ingredients } = req.body;
  if (!name || !recipeYield || !Array.isArray(ingredients) || ingredients.length === 0) {
    return res.status(400).json({ error: 'Неполные данные' });
  }

  try {
    const recipesData = await getSheetData('РЕЦЕПТЫ!A:A');
    let lastId = 0;
    recipesData.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
    const newRecipeId = lastId + 1;

    await appendRow('РЕЦЕПТЫ!A:E', [newRecipeId, name, recipeYield, '', '']);

    const ingData = await getSheetData('ИНГРЕДИЕНТЫ!A:E');
    const ingByName = new Map();
    let maxIngId = 0;
    ingData.forEach(ing => {
      const id = parseInt(ing[0]);
      ingByName.set(ing[1], id);
      if (id > maxIngId) maxIngId = id;
    });

    const compositionRows = [];
    const newIngredients = [];

    for (const ing of ingredients) {
      let ingId = ingByName.get(ing.name);
      if (!ingId) {
        maxIngId++;
        ingId = maxIngId;
        const defaultUnit = 'г';
        const defaultPrice = 0;
        await appendRow('ИНГРЕДИЕНТЫ!A:E', [ingId, ing.name, 'Автосоздан', defaultUnit, defaultPrice]);
        await appendRow('СКЛАД!A:D', [ingId, ing.name, 0, 0]);
        ingByName.set(ing.name, ingId);
        newIngredients.push(ing.name);
      }
      compositionRows.push([newRecipeId, ingId, ing.amountG]);
    }

    if (compositionRows.length > 0) {
      for (const row of compositionRows) {
        await appendRow('СОСТАВ_РЕЦЕПТА!A:C', row);
      }
    }

    let message = `Рецепт "${name}" добавлен с ID ${newRecipeId}.`;
    if (newIngredients.length > 0) {
      message += `\n➕ Автоматически добавлены ингредиенты: ${newIngredients.join(', ')} (цена 0, ед.изм. "г")`;
    }
    res.json({ success: true, message });
  } catch (err) {
    console.error('Ошибка создания рецепта:', err);
    res.status(500).json({ error: 'Ошибка сервера' });
  }
});

// ---------- ОПЕРАЦИИ СО СКЛАДОМ ----------
app.post('/api/stock/transaction', authenticateToken, async (req, res) => {
  const { materialName, operation, quantity, comment } = req.body;
  if (!materialName || !operation || !quantity || quantity <= 0) {
    return res.status(400).json({ error: 'Неверные параметры' });
  }
  try {
    const stockData = await getSheetData('СКЛАД!A:D');
    let rowIndex = -1;
    let currentStock = 0;
    for (let i = 0; i < stockData.length; i++) {
      if (stockData[i][1] === materialName) {
        rowIndex = i + 2;
        currentStock = parseFloat(stockData[i][3]) || 0;
        break;
      }
    }
    if (rowIndex === -1) {
      return res.status(404).json({ error: 'Материал не найден' });
    }

    let newStock;
    if (operation === 'Приход') {
      newStock = currentStock + quantity;
    } else if (operation === 'Расход') {
      if (currentStock < quantity) {
        return res.status(400).json({ error: `Недостаточно материала: остаток ${currentStock}` });
      }
      newStock = currentStock - quantity;
    } else {
      return res.status(400).json({ error: 'Операция должна быть "Приход" или "Расход"' });
    }

    await updateRow('СКЛАД', rowIndex, [stockData[rowIndex-2][0], stockData[rowIndex-2][1], stockData[rowIndex-2][2], newStock]);

    const now = new Date().toISOString();
    await appendRow('ИСТОРИЯ!A:F', [now, materialName, operation, quantity, newStock, comment || '']);

    res.json({ success: true, message: `${operation} ${quantity} → новый остаток ${newStock}` });
  } catch (err) {
    console.error('Ошибка операции со складом:', err);
    res.status(500).json({ error: 'Внутренняя ошибка' });
  }
});

app.post('/api/stock/batch', authenticateToken, async (req, res) => {
  const operations = req.body;
  if (!Array.isArray(operations) || operations.length === 0) {
    return res.status(400).json({ error: 'Нет операций' });
  }

  const errors = [];
  const successOps = [];

  for (const op of operations) {
    try {
      const { material, operation, quantity, comment } = op;
      if (!material || !operation || !quantity || quantity <= 0) {
        errors.push(`Неверные данные для ${material}`);
        continue;
      }
      const stockData = await getSheetData('СКЛАД!A:D');
      let rowIndex = -1;
      let currentStock = 0;
      for (let i = 0; i < stockData.length; i++) {
        if (stockData[i][1] === material) {
          rowIndex = i + 2;
          currentStock = parseFloat(stockData[i][3]) || 0;
          break;
        }
      }
      if (rowIndex === -1) {
        errors.push(`Материал "${material}" не найден`);
        continue;
      }
      let newStock;
      if (operation === 'Приход') {
        newStock = currentStock + quantity;
      } else if (operation === 'Расход') {
        if (currentStock < quantity) {
          errors.push(`Недостаточно "${material}": остаток ${currentStock}`);
          continue;
        }
        newStock = currentStock - quantity;
      } else {
        errors.push(`Некорректная операция для "${material}"`);
        continue;
      }
      await updateRow('СКЛАД', rowIndex, [stockData[rowIndex-2][0], stockData[rowIndex-2][1], stockData[rowIndex-2][2], newStock]);
      const now = new Date().toISOString();
      await appendRow('ИСТОРИЯ!A:F', [now, material, operation, quantity, newStock, comment || '']);
      successOps.push(material);
    } catch (err) {
      errors.push(`Ошибка при обработке ${op.material}: ${err.message}`);
    }
  }

  if (errors.length > 0) {
    return res.status(400).json({ success: false, errors });
  }
  res.json({ success: true, message: `Выполнено ${successOps.length} операций.` });
});

// ---------- СПИСАНИЕ ИНГРЕДИЕНТОВ (калькулятор) ----------
app.post('/api/stock/writeoff', authenticateToken, async (req, res) => {
  const { items, comment } = req.body;
  if (!items || !Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'Нет ингредиентов для списания' });
  }
  try {
    const stockData = await getSheetData('СКЛАД!A2:D');
    const stockMap = new Map();
    stockData.forEach(row => {
      const name = row[1];
      const stock = parseFloat(row[3]) || 0;
      const rowIndex = stockData.indexOf(row) + 2;
      stockMap.set(name, { stock, rowIndex });
    });

    const errors = [];
    const transactions = [];

    for (const item of items) {
      const stockInfo = stockMap.get(item.name);
      if (!stockInfo) {
        errors.push(`Ингредиент "${item.name}" не найден`);
        continue;
      }
      if (stockInfo.stock < item.neededGrams) {
        errors.push(`Недостаточно "${item.name}" (остаток: ${stockInfo.stock}, нужно: ${item.neededGrams})`);
        continue;
      }
      const newStock = stockInfo.stock - item.neededGrams;
      transactions.push({
        name: item.name,
        qty: item.neededGrams,
        rowIndex: stockInfo.rowIndex,
        newStock
      });
    }

    if (errors.length > 0) {
      return res.status(400).json({ error: errors.join('; ') });
    }

    const now = new Date().toISOString();
    for (const t of transactions) {
      await updateRow('СКЛАД', t.rowIndex, [t.rowIndex, t.name, '', t.newStock]);
      await appendRow('ИСТОРИЯ!A:F', [now, t.name, 'Расход', t.qty, t.newStock, comment || 'Списание через калькулятор']);
    }

    res.json({ success: true, message: `Списано ${transactions.length} позиций` });
  } catch (err) {
    console.error('Ошибка списания:', err);
    res.status(500).json({ error: 'Внутренняя ошибка сервера' });
  }
});

// ---------- КАТЕГОРИИ ----------
app.get('/api/categories', async (req, res) => {
  const rows = await getSheetData('КАТЕГОРИИ!A2:C');
  const categories = rows.map(row => ({ id: parseInt(row[0]), name: row[1], sortOrder: parseInt(row[2]) || 0 }));
  res.json(categories);
});

app.post('/api/categories', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Доступ запрещён' });
  const { name } = req.body;
  if (!name) return res.status(400).json({ error: 'Название обязательно' });
  const rows = await getSheetData('КАТЕГОРИИ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  await appendRow('КАТЕГОРИИ!A:C', [newId, name, 0]);
  res.json({ success: true, id: newId });
});

app.put('/api/categories/:id', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Доступ запрещён' });
  const catId = parseInt(req.params.id);
  const { name } = req.body;
  const rows = await getSheetData('КАТЕГОРИИ!A2:C');
  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === catId) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Категория не найдена' });
  await updateRow('КАТЕГОРИИ', rowIndex, [catId, name, rows[rowIndex-2][2] || 0]);
  res.json({ success: true });
});

app.delete('/api/categories/:id', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Доступ запрещён' });
  const catId = parseInt(req.params.id);
  const rows = await getSheetData('КАТЕГОРИИ!A:A');
  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === catId) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Категория не найдена' });
  await deleteRow('КАТЕГОРИИ', rowIndex);
  res.json({ success: true });
});

// ---------- ТОВАРЫ ----------
app.get('/api/products', async (req, res) => {
  const rows = await getSheetData('ТОВАРЫ!A2:G');
  const products = rows.map(row => ({
    id: parseInt(row[0]),
    categoryId: parseInt(row[1]),
    name: row[2],
    price: parsePrice(row[3]),
    description: row[4] || '',
    isActive: row[5] === 'TRUE' || row[5] === true,
    createdAt: row[6]
  }));
  res.json(products);
});

app.post('/api/products', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin' && req.user.role !== 'manager') {
    return res.status(403).json({ error: 'Доступ запрещён' });
  }
  const { categoryId, name, price, description } = req.body;
  if (!categoryId || !name || !price) return res.status(400).json({ error: 'Не хватает данных' });
  const rows = await getSheetData('ТОВАРЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const now = new Date().toISOString();
  await appendRow('ТОВАРЫ!A:G', [newId, categoryId, name, price, description || '', true, now]);
  res.json({ success: true, id: newId });
});

app.put('/api/products/:id', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Доступ запрещён' });
  const productId = parseInt(req.params.id);
  const { categoryId, name, price, description, isActive } = req.body;
  const rows = await getSheetData('ТОВАРЫ!A2:G');
  let rowIndex = -1;
  let oldRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === productId) {
      rowIndex = i + 2;
      oldRow = rows[i];
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Товар не найден' });
  const newRow = [
    productId,
    categoryId !== undefined ? categoryId : oldRow[1],
    name !== undefined ? name : oldRow[2],
    price !== undefined ? price : oldRow[3],
    description !== undefined ? description : oldRow[4],
    isActive !== undefined ? isActive : oldRow[5],
    oldRow[6]
  ];
  const success = await updateRow('ТОВАРЫ', rowIndex, newRow);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка обновления товара' });
});

app.delete('/api/products/:id', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Доступ запрещён' });
  const productId = parseInt(req.params.id);
  const rows = await getSheetData('ТОВАРЫ!A:A');
  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === productId) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Товар не найден' });
  const success = await deleteRow('ТОВАРЫ', rowIndex);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка удаления товара' });
});

// ---------- ТОВАРЫ ЗАКАЗА ----------
app.post('/api/order-products', authenticateToken, async (req, res) => {
  const { orderId, productId, quantity, price } = req.body;
  await appendRow('ЗАКАЗЫ_ТОВАРЫ!A:D', [orderId, productId, quantity, price]);
  res.json({ success: true });
});

app.get('/api/order-products/:orderId', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.orderId);
  const rows = await getSheetData('ЗАКАЗЫ_ТОВАРЫ!A2:D');
  const products = rows.filter(row => parseInt(row[0]) === orderId).map(row => ({
    productId: parseInt(row[1]),
    quantity: parseFloat(row[2]),
    price: parsePrice(row[3])
  }));
  res.json(products);
});

app.delete('/api/order-products/:orderId', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.orderId);
  const rows = await getSheetData('ЗАКАЗЫ_ТОВАРЫ!A2:D');
  const rowsToDelete = [];
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === orderId) rowsToDelete.push(i + 2);
  }
  for (const row of rowsToDelete.reverse()) {
    await deleteRow('ЗАКАЗЫ_ТОВАРЫ', row);
  }
  res.json({ success: true });
});

// ---------- ФИНАНСЫ (редактирование и удаление) ----------
app.put('/api/transactions/:id', authenticateToken, async (req, res) => {
  const transId = parseInt(req.params.id);
  const { date, type, category, amount, comment } = req.body;
  const rows = await getSheetData('ФИНАНСЫ!A2:G');
  let rowIndex = -1;
  let oldRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === transId) {
      rowIndex = i + 2;
      oldRow = rows[i];
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Транзакция не найдена' });

  const newRow = [
    transId,
    date ? new Date(date).toISOString() : oldRow[1],
    type || oldRow[2],
    category || oldRow[3],
    amount !== undefined ? amount : oldRow[4],
    comment || oldRow[5],
    oldRow[6] || ''
  ];
  const success = await updateRow('ФИНАНСЫ', rowIndex, newRow);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка обновления транзакции' });
});

app.delete('/api/transactions/:id', authenticateToken, async (req, res) => {
  const transId = parseInt(req.params.id);
  console.log(`🗑️ Запрос на удаление транзакции ${transId}`);

  // Читаем все строки листа ФИНАНСЫ
  const rows = await getSheetData('ФИНАНСЫ!A2:G');
  let rowIndex = -1;
  let foundRow = null;
  for (let i = 0; i < rows.length; i++) {
    const rowId = parseInt(rows[i][0]);
    if (rowId === transId) {
      rowIndex = i + 2; // +2 потому что строка 1 – заголовок, i начинается с 0
      foundRow = rows[i];
      break;
    }
  }

  if (rowIndex === -1) {
    console.log(`❌ Транзакция ${transId} не найдена в таблице`);
    return res.status(404).json({ error: 'Транзакция не найдена' });
  }

  console.log(`✅ Найдена строка ${rowIndex}:`, foundRow);

  // Пытаемся удалить
  const success = await deleteRow('ФИНАНСЫ', rowIndex);
  if (success) {
    console.log(`✅ Транзакция ${transId} успешно удалена`);
    res.json({ success: true });
  } else {
    console.log(`❌ Ошибка удаления транзакции ${transId}`);
    res.status(500).json({ error: 'Ошибка удаления транзакции' });
  }
});

// ---------- ИСТОРИЯ ЗАКАЗОВ ----------
app.get('/api/order-history-full/:orderId', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.orderId);
  const rows = await getSheetData('ИСТОРИЯ_ЗАКАЗОВ_СНАПШОТЫ!A2:K');
  const history = rows
    .filter(row => parseInt(row[0]) === orderId)
    .map(row => {
      let products = [];
      try {
        products = JSON.parse(row[10] || '[]');
      } catch(e) {}
      return {
        changedAt: row[1],
        userName: row[9] || `Пользователь ${row[2]}`,
        clientId: row[3],
        price: parseFloat(row[4]),
        status: row[5],
        details: row[6],
        delivery: row[7],
        executionDate: row[8],
        userId: parseInt(row[2]),
        products: products
      };
    })
    .sort((a,b) => new Date(b.changedAt) - new Date(a.changedAt));
  res.json(history);
});

// ---------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------
function parsePrice(value) {
  if (value === undefined || value === null) return 0;
  if (typeof value === 'number') return value;
  let str = String(value).trim().replace(',', '.');
  let num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

async function findOrCreateClient(name, phone, address) {
  console.log('findOrCreateClient args:', { name, phone, address });
  const hasName = name && name.trim() !== '' && name.trim() !== 'Аноним';
  const hasPhone = phone && phone.trim() !== '' && phone.trim() !== 'не указан';
  const hasAddress = address && address.trim() !== '';

  if (!hasName && !hasPhone && !hasAddress) return null;

  let client = null;
  let clientRowIndex = -1;

  if (hasPhone) {
    const normalizedPhone = phone.replace(/[^0-9+]/g, '');
    const clients = await getSheetData('КЛИЕНТЫ!A2:E');
    for (let i = 0; i < clients.length; i++) {
      const c = clients[i];
      const cPhone = c[2] ? c[2].replace(/[^0-9+]/g, '') : '';
      if (cPhone === normalizedPhone) {
        client = c;
        clientRowIndex = i + 2;
        break;
      }
    }
  }

  if (client) {
    const clientId = Number(client[0]);
    let needUpdate = false;
    if (hasName && (!client[1] || client[1] === 'Аноним')) {
      client[1] = name;
      needUpdate = true;
    }
    if (hasPhone && (!client[2] || client[2] === 'не указан')) {
      client[2] = phone.replace(/[^0-9+]/g, '');
      needUpdate = true;
    }
    if (hasAddress && !client[3]) {
      client[3] = address;
      needUpdate = true;
    }
    if (needUpdate && clientRowIndex !== -1) {
      await updateRow('КЛИЕНТЫ', clientRowIndex, [clientId, client[1], client[2], client[3], client[4]]);
    }
    return clientId;
  } else {
    const clients = await getSheetData('КЛИЕНТЫ!A2:E');
    const ids = clients.map(c => Number(c[0])).filter(id => !isNaN(id));
    const newId = ids.length ? Math.max(...ids) + 1 : 1;
    await appendRow('КЛИЕНТЫ!A:E', [
      newId,
      hasName ? name : '',
      hasPhone ? phone.replace(/[^0-9+]/g, '') : '',
      hasAddress ? address : '',
      ''
    ]);
    return newId;
  }
}

async function getOrderProductsWithNames(orderId) {
  const productsRows = await getSheetData('ЗАКАЗЫ_ТОВАРЫ!A2:D');
  const orderProducts = productsRows.filter(row => parseInt(row[0]) === orderId);
  const allProducts = await getSheetData('ТОВАРЫ!A2:G');
  const productMap = new Map();
  allProducts.forEach(p => productMap.set(parseInt(p[0]), p[2]));
  const result = [];
  for (const op of orderProducts) {
    const productId = parseInt(op[1]);
    let name = productMap.get(productId);
    if (!name) name = `Товар #${productId}`;
    result.push({
      name: name,
      quantity: parseFloat(op[2]),
      price: parseFloat(op[3])
    });
  }
  return result;
}

async function saveOrderSnapshot(orderId, userId, userName, products) {
  const rows = await getSheetData('ЗАКАЗЫ!A2:I');
  let orderRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === orderId) {
      orderRow = rows[i];
      break;
    }
  }
  if (!orderRow) return;

  const now = new Date().toISOString();
  const productsJson = JSON.stringify(products.map(p => ({
    name: p.name,
    quantity: p.quantity,
    price: p.price
  })));

  await appendRow('ИСТОРИЯ_ЗАКАЗОВ_СНАПШОТЫ!A:K', [
    orderId,
    now,
    userId,
    orderRow[2],
    orderRow[3],
    orderRow[4],
    orderRow[5],
    orderRow[6],
    orderRow[7],
    userName,
    productsJson
  ]);
}

// ---------- ЕДИНСТВЕННАЯ ФУНКЦИЯ СОЗДАНИЯ ТРАНЗАКЦИИ ----------
async function createTransactionForOrder(orderId, price, clientId) {
  const financeRows = await getSheetData('ФИНАНСЫ!A:A');
  let lastId = 0;
  financeRows.forEach(row => {
    const id = parseInt(row[0]);
    if (!isNaN(id) && id > lastId) lastId = id;
  });
  const newId = lastId + 1;
  const now = new Date().toISOString();
  const comment = `Оплата заказа #${orderId}`;

  const success = await appendRow('ФИНАНСЫ!A:G', [
    newId, now, 'Доход', 'Заказ', price, comment, orderId
  ]);
  if (!success) console.error(`Не удалось создать транзакцию для заказа ${orderId}`);
}

// ---------- ЕДИНСТВЕННАЯ ФУНКЦИЯ УДАЛЕНИЯ ТРАНЗАКЦИИ ----------
async function deleteTransactionByOrderId(orderId) {
  const rows = await getSheetData('ФИНАНСЫ!A2:G');
  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][6]) === orderId) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) {
    console.log(`Транзакция для заказа ${orderId} не найдена, удаление не требуется`);
    return;
  }
  const success = await deleteRow('ФИНАНСЫ', rowIndex);
  if (!success) console.error(`Ошибка удаления транзакции для заказа ${orderId}`);
}

// ---------- УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ ----------
function generateSecurePassword() {
  const length = 10;
  const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()";
  let password = "";
  for (let i = 0; i < length; i++) {
    const randomIndex = Math.floor(Math.random() * charset.length);
    password += charset[randomIndex];
  }
  return password;
}

function deriveKeyAndHash(password, salt) {
  const combined = password + salt;
  const digest = crypto.createHash('sha256').update(combined).digest('hex');
  return digest;
}

app.get('/api/users', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') {
    return res.status(403).json({ error: 'Доступ запрещён' });
  }
  try {
    const usersData = await getSheetData('ПОЛЬЗОВАТЕЛИ!A2:H');
    const users = usersData.map(row => ({
      id: parseInt(row[0]),
      name: row[1],
      email: row[2],
      role: row[3],
    }));
    res.json(users);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Ошибка загрузки пользователей' });
  }
});

app.post('/api/users', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') {
    return res.status(403).json({ error: 'Доступ запрещён' });
  }
  const { name, email, role } = req.body;
  if (!name || !email || !role) {
    return res.status(400).json({ error: 'Не хватает данных' });
  }
  try {
    const users = await getSheetData('ПОЛЬЗОВАТЕЛИ!A:C');
    if (users.some(u => u[2] === email)) {
      return res.status(400).json({ error: 'Пользователь с таким email уже существует' });
    }
    const plainPassword = generateSecurePassword();
    const salt = generateSecurePassword();
    const passwordHash = deriveKeyAndHash(plainPassword, salt);
    let lastId = 0;
    users.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
    const newId = lastId + 1;
    await appendRow('ПОЛЬЗОВАТЕЛИ!A:H', [newId, name, email, role, passwordHash, salt, '', '[]']);
    res.json({ success: true, password: plainPassword });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Ошибка создания пользователя' });
  }
});

app.delete('/api/users/:id', authenticateToken, async (req, res) => {
  if (req.user.role !== 'admin') {
    return res.status(403).json({ error: 'Доступ запрещён' });
  }
  const userId = parseInt(req.params.id);
  try {
    const usersData = await getSheetData('ПОЛЬЗОВАТЕЛИ!A:A');
    let rowIndex = -1;
    for (let i = 0; i < usersData.length; i++) {
      if (parseInt(usersData[i][0]) === userId) {
        rowIndex = i + 2;
        break;
      }
    }
    if (rowIndex === -1) {
      return res.status(404).json({ error: 'Пользователь не найден' });
    }
    await deleteRow('ПОЛЬЗОВАТЕЛИ', rowIndex);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Ошибка удаления пользователя' });
  }
});

app.post('/api/ingredients', authenticateToken, async (req, res) => {
  const { name, category, unit, price } = req.body;
  if (!name || !unit || price === undefined) {
    return res.status(400).json({ error: 'Не хватает данных' });
  }
  try {
    const ingData = await getSheetData('ИНГРЕДИЕНТЫ!A:E');
    if (ingData.some(row => row[1] === name)) {
      return res.status(400).json({ error: 'Ингредиент уже существует' });
    }
    let lastId = 0;
    ingData.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
    const newId = lastId + 1;
    await appendRow('ИНГРЕДИЕНТЫ!A:E', [newId, name, category || '', unit, price]);
    await appendRow('СКЛАД!A:D', [newId, name, 0, 0]);
    res.json({ success: true, message: `Ингредиент "${name}" добавлен` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Ошибка добавления ингредиента' });
  }
});

// ---------- ЗАПУСК СЕРВЕРА ----------
process.on('uncaughtException', (err) => {
  console.error('❌ Uncaught Exception:', err);
});
process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ Unhandled Rejection:', reason);
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
