require('dotenv').config();
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
  const token = req.cookies?.access_token;
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
    price: row[3],
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
    date: row[1],
    type: row[2],
    category: row[3],
    amount: parseFloat(row[4]),
    comment: row[5],
  }));
  if (startDate && endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    transactions = transactions.filter(t => {
      const d = new Date(t.date);
      return d >= start && d <= end;
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
  res.cookie('access_token', token, { httpOnly: true, maxAge: 30 * 24 * 60 * 60 * 1000, sameSite: 'lax' });
  res.json({ success: true, name: user[1], role: user[3], email });
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
