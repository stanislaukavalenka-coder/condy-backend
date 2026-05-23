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
app.use(cors({ origin: true, credentials: true }));
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
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${range}!A${rowNumber}:${String.fromCharCode(64 + values.length)}${rowNumber}`,
      valueInputOption: 'RAW',
      resource: { values: [values] },
    });
    return true;
  } catch (err) {
    console.error(`Ошибка обновления строки ${rowNumber}:`, err.message);
    return false;
  }
}

async function deleteRow(sheetName, rowNumber) {
  const sheetId = await getSheetId(sheetName);
  if (!sheetId) return false;
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
    return true;
  } catch (err) {
    console.error(`Ошибка удаления строки ${rowNumber}:`, err.message);
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
    price: parseFloat(row[3]) || 0,
    status: row[4],
    details: row[5],
    delivery: row[6],
    executionDate: row[7],
  }));
  res.json(orders);
});

app.post('/api/orders', authenticateToken, async (req, res) => {
  const { clientId, price, status, details, delivery, executionDate } = req.body;
  const rows = await getSheetData('ЗАКАЗЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const now = new Date().toISOString();
  const success = await appendRow('ЗАКАЗЫ!A:I', [
    newId, now, clientId, price, status, details, delivery, executionDate, req.user.userId
  ]);
  if (success) res.json({ success: true, orderId: newId });
  else res.status(500).json({ error: 'Ошибка создания заказа' });
});

app.put('/api/orders/:id', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.id);
  const updates = req.body;
  const rows = await getSheetData('ЗАКАЗЫ!A:I');
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
  const newRow = [
    oldRow[0],
    oldRow[1],
    updates.clientId !== undefined ? updates.clientId : oldRow[2],
    updates.price !== undefined ? updates.price : oldRow[3],
    updates.status !== undefined ? updates.status : oldRow[4],
    updates.details !== undefined ? updates.details : oldRow[5],
    updates.delivery !== undefined ? updates.delivery : oldRow[6],
    updates.executionDate !== undefined ? updates.executionDate : oldRow[7],
    req.user.userId,
  ];
  const success = await updateRow('ЗАКАЗЫ', rowIndex, newRow);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка обновления заказа' });
});

app.delete('/api/orders/:id', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.id);
  const rows = await getSheetData('ЗАКАЗЫ!A:A');
  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === orderId) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Заказ не найден' });
  const success = await deleteRow('ЗАКАЗЫ', rowIndex);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка удаления заказа' });
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
  const rows = await getSheetData('КЛИЕНТЫ!A:E');
  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i][0]) === clientId) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) return res.status(404).json({ error: 'Клиент не найден' });
  const newRow = [clientId, name, phone || '', address || '', notes || ''];
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
  // реализация обновления склада (приход/расход) – вы можете добавить позже
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
  const rows = await getSheetData('ФИНАНСЫ!A2:F');
  let transactions = rows.map(row => ({
    id: row[0],
    let rawDate = row[1];
let isoDate = null;
if (typeof rawDate === 'string') {
  const parts = rawDate.split(' ');
  const dateParts = parts[0].split('.');
  if (dateParts.length === 3) {
    const year = parseInt(dateParts[2]);
    const month = parseInt(dateParts[1]) - 1;
    const day = parseInt(dateParts[0]);
    let hour = 0, minute = 0, second = 0;
    if (parts[1]) {
      const timeParts = parts[1].split(':');
      hour = parseInt(timeParts[0]);
      minute = parseInt(timeParts[1]);
      second = parseInt(timeParts[2] || 0);
    }
    const d = new Date(year, month, day, hour, minute, second);
    isoDate = d.toISOString();
  } else {
    isoDate = rawDate;
  }
} else if (rawDate instanceof Date) {
  isoDate = rawDate.toISOString();
} else {
  isoDate = rawDate;
}
    type: row[2],
    category: row[3],
    amount: parseFloat(row[4]) || 0,
    comment: row[5],
  }));
  if (startDate && endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    start.setHours(0,0,0,0);
    end.setHours(23,59,59,999);
    transactions = transactions.filter(t => {
      let d;
      if (typeof t.date === 'string') {
        const parts = t.date.split(' ');
        const dateParts = parts[0].split('.');
        const timeParts = parts[1] ? parts[1].split(':') : [0,0,0];
        if (dateParts.length === 3) {
          d = new Date(
            parseInt(dateParts[2]),
            parseInt(dateParts[1]) - 1,
            parseInt(dateParts[0]),
            parseInt(timeParts[0]),
            parseInt(timeParts[1]),
            parseInt(timeParts[2] || 0)
          );
        } else {
          d = new Date(t.date);
        }
      } else if (t.date instanceof Date) {
        d = t.date;
      } else {
        d = new Date(t.date);
      }
      return !isNaN(d.getTime()) && d >= start && d <= end;
    });
  }
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
  const { type, category, amount, comment } = req.body;
  const rows = await getSheetData('ФИНАНСЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const now = new Date().toISOString();
  const success = await appendRow('ФИНАНСЫ!A:F', [newId, now, type, category, amount, comment]);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка добавления транзакции' });
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
  const token = jwt.sign({ userId: user[0], name: user[1], email, role: user[3] }, JWT_SECRET, { expiresIn: '30d' });
  // Не устанавливаем httpOnly куку, а возвращаем токен в теле
  res.json({ success: true, name: user[1], role: user[3], email, token });
});

app.get('/api/me', authenticateToken, (req, res) => {
  res.json(req.user);
});

app.post('/api/logout', (req, res) => {
  res.clearCookie('access_token');
  res.json({ success: true });
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});

app.post('/api/recipes/calculate', async (req, res) => {
  const { recipeId, desiredWeight } = req.body;
  if (!recipeId || !desiredWeight) {
    return res.status(400).json({ error: 'recipeId и desiredWeight обязательны' });
  }
  try {
    // Получаем рецепт (выход)
    const recipes = await getSheetData('РЕЦЕПТЫ!A2:D');
    const recipe = recipes.find(r => parseInt(r[0]) === recipeId);
    if (!recipe) return res.status(404).json({ error: 'Рецепт не найден' });
    const recipeYield = parseFloat(recipe[2]); // столбец C (выход в граммах)
    if (!recipeYield) return res.status(400).json({ error: 'В рецепте не указан выход' });

    const coefficient = desiredWeight / recipeYield;

    // Получаем состав рецепта
    const composition = await getSheetData('СОСТАВ_РЕЦЕПТА!A:C');
    const recipeIngredients = composition.filter(row => parseInt(row[0]) === recipeId);
    if (recipeIngredients.length === 0) {
      return res.status(404).json({ error: 'В рецепте нет ингредиентов' });
    }

    // Получаем список ингредиентов с названиями
    const ingredients = await getSheetData('ИНГРЕДИЕНТЫ!A:E');
    const ingMap = new Map();
    ingredients.forEach(ing => ingMap.set(parseInt(ing[0]), { name: ing[1], unit: ing[3], price: parseFloat(ing[4]) || 0 }));

    const resultIngredients = [];
    for (const comp of recipeIngredients) {
      const ingId = parseInt(comp[1]);
      const amount = parseFloat(comp[2]); // количество в граммах на одну порцию/вес рецепта
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
    // 1. Получить следующий ID для рецепта
    const recipesData = await getSheetData('РЕЦЕПТЫ!A:A');
    let lastId = 0;
    recipesData.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
    const newRecipeId = lastId + 1;

    // 2. Добавить рецепт в лист РЕЦЕПТЫ
    await appendRow('РЕЦЕПТЫ!A:E', [newRecipeId, name, recipeYield, '', '']);

    // 3. Получить список ингредиентов (справочник)
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
        // Создаём новый ингредиент
        maxIngId++;
        ingId = maxIngId;
        const defaultUnit = 'г';
        const defaultPrice = 0;
        await appendRow('ИНГРЕДИЕНТЫ!A:E', [ingId, ing.name, 'Автосоздан', defaultUnit, defaultPrice]);
        // Также добавляем в СКЛАД (если есть лист СКЛАД)
        await appendRow('СКЛАД!A:D', [ingId, ing.name, 0, 0]);
        ingByName.set(ing.name, ingId);
        newIngredients.push(ing.name);
      }
      compositionRows.push([newRecipeId, ingId, ing.amountG]);
    }

    // 4. Добавить состав рецепта
    if (compositionRows.length > 0) {
      const lastRow = await getSheetData('СОСТАВ_РЕЦЕПТА!A:A');
      const startRow = lastRow.length + 2; // +2 потому что данные с A2
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

app.post('/api/stock/transaction', authenticateToken, async (req, res) => {
  const { materialName, operation, quantity, comment } = req.body;
  if (!materialName || !operation || !quantity || quantity <= 0) {
    return res.status(400).json({ error: 'Неверные параметры' });
  }
  try {
    // Найти строку материала в СКЛАД
    const stockData = await getSheetData('СКЛАД!A:D');
    let rowIndex = -1;
    let currentStock = 0;
    for (let i = 0; i < stockData.length; i++) {
      if (stockData[i][1] === materialName) {
        rowIndex = i + 2; // +2 потому что данные с A2
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

    // Обновить остаток
    const range = `СКЛАД!D${rowIndex}`;
    await updateRow('СКЛАД', rowIndex, [stockData[rowIndex-2][0], stockData[rowIndex-2][1], stockData[rowIndex-2][2], newStock]);

    // Добавить запись в ИСТОРИЯ
    const now = new Date().toISOString();
    await appendRow('ИСТОРИЯ!A:F', [now, materialName, operation, quantity, newStock, comment || '']);

    res.json({ success: true, message: `${operation} ${quantity} → новый остаток ${newStock}` });
  } catch (err) {
    console.error('Ошибка операции со складом:', err);
    res.status(500).json({ error: 'Внутренняя ошибка' });
  }
});

app.post('/api/stock/batch', authenticateToken, async (req, res) => {
  const operations = req.body; // ожидаем массив { material, operation, quantity, comment }
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
      // Найти строку материала в СКЛАД
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
      // Обновление остатка
      await updateRow('СКЛАД', rowIndex, [stockData[rowIndex-2][0], stockData[rowIndex-2][1], stockData[rowIndex-2][2], newStock]);
      // Запись в историю
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

// Функции для работы с паролями (аналогично вашим GAS)
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
  // Только админ может просматривать
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
    // Проверка на существование
    const users = await getSheetData('ПОЛЬЗОВАТЕЛИ!A:C');
    if (users.some(u => u[2] === email)) {
      return res.status(400).json({ error: 'Пользователь с таким email уже существует' });
    }
    const plainPassword = generateSecurePassword();
    const salt = generateSecurePassword();
    const passwordHash = deriveKeyAndHash(plainPassword, salt);
    // Получить новый ID
    let lastId = 0;
    users.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
    const newId = lastId + 1;
    // Добавить строку
    await appendRow('ПОЛЬЗОВАТЕЛИ!A:H', [newId, name, email, role, passwordHash, salt, '', '[]']);
    // Отправить пароль по email (опционально)
    // Здесь можно добавить отправку email, например, через nodemailer или другой сервис
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
    // Проверка на дубликат
    if (ingData.some(row => row[1] === name)) {
      return res.status(400).json({ error: 'Ингредиент уже существует' });
    }
    let lastId = 0;
    ingData.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
    const newId = lastId + 1;
    await appendRow('ИНГРЕДИЕНТЫ!A:E', [newId, name, category || '', unit, price]);
    // Также добавить в СКЛАД
    await appendRow('СКЛАД!A:D', [newId, name, 0, 0]);
    res.json({ success: true, message: `Ингредиент "${name}" добавлен` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Ошибка добавления ингредиента' });
  }
});


