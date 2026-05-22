require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const fs = require('fs');

const app = express();
app.use(express.json());

// CORS (разрешаем всё для теста)
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('Access-Control-Allow-Credentials', 'true');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET) {
  console.error('❌ JWT_SECRET не задан');
  process.exit(1);
}

// Загрузка ключа сервисного аккаунта
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

// Вспомогательные функции
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
    console.error(`Ошибка добавления в ${range}:`, err.message);
    return false;
  }
}
async function updateRow(range, rowNumber, values) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${range}!A${rowNumber}:${String.fromCharCode(64 + values.length)}${rowNumber}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
    return true;
  } catch (err) {
    console.error(`Ошибка обновления строки ${rowNumber}:`, err.message);
    return false;
  }
}
async function deleteRow(sheetName, rowNumber) {
  try {
    const sheetId = await getSheetId(sheetName);
    if (!sheetId) return false;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: {
        requests: [{
          deleteDimension: {
            range: { sheetId, dimension: 'ROWS', startIndex: rowNumber - 1, endIndex: rowNumber }
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

// Middleware проверки JWT
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

// ========== МАРШРУТЫ ==========
app.get('/ping', (req, res) => res.json({ message: 'pong' }));

// Заказы (чтение)
app.get('/api/orders', async (req, res) => {
  const rows = await getSheetData('ЗАКАЗЫ!A2:H');
  const orders = rows.map(row => ({
    id: parseInt(row[0]),
    created: row[1],
    clientId: parseInt(row[2]),
    price: parseFloat(row[3]),
    status: row[4],
    details: row[5],
    delivery: row[6],
    executionDate: row[7],
  }));
  res.json(orders);
});
app.post('/api/orders', authenticateToken, async (req, res) => {
  const { clientId, price, status, details, delivery, executionDate } = req.body;
  if (!clientId || price === undefined || !executionDate) return res.status(400).json({ error: 'Не хватает полей' });
  const rows = await getSheetData('ЗАКАЗЫ!A:A');
  let lastId = 0;
  rows.forEach(r => { if (r[0]) lastId = Math.max(lastId, parseInt(r[0])); });
  const newId = lastId + 1;
  const now = new Date().toISOString();
  const success = await appendRow('ЗАКАЗЫ!A:H', [newId, now, clientId, price, status || 'в работе', details || '', delivery || '', executionDate]);
  if (success) res.json({ success: true, orderId: newId });
  else res.status(500).json({ error: 'Ошибка создания' });
});
app.put('/api/orders/:id', authenticateToken, async (req, res) => {
  const orderId = parseInt(req.params.id);
  const updates = req.body;
  const rows = await getSheetData('ЗАКАЗЫ!A:H');
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
    updates.executionDate !== undefined ? updates.executionDate : oldRow[7]
  ];
  const success = await updateRow('ЗАКАЗЫ', rowIndex, newRow);
  if (success) res.json({ success: true });
  else res.status(500).json({ error: 'Ошибка обновления' });
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
  else res.status(500).json({ error: 'Ошибка удаления' });
});

// Клиенты
app.get('/api/clients', async (req, res) => {
  const rows = await getSheetData('КЛИЕНТЫ!A2:E');
  const clients = rows.map(row => ({ id: parseInt(row[0]), name: row[1], phone: row[2], address: row[3], notes: row[4] }));
  res.json(clients);
});
app.post('/api/clients', authenticateToken, async (req, res) => {
  const { name, phone, address, notes } = req.body;
  if (!name || !phone) return res.status(400).json({ error: 'Имя и телефон обязательны' });
  const rows = await getSheetData('КЛИЕНТЫ!A:A');
  let lastId = 0;
  rows.forEach(r => { if (r[0]) lastId = Math.max(lastId, parseInt(r[0])); });
  const newId = lastId + 1;
  const success = await appendRow('КЛИЕНТЫ!A:E', [newId, name, phone, address || '', notes || '']);
  if (success) res.json({ success: true, id: newId });
  else res.status(500).json({ error: 'Ошибка создания' });
});

// Склад
app.get('/api/stock', async (req, res) => {
  const rows = await getSheetData('СКЛАД!A2:D');
  const stock = rows.map(row => ({ id: parseInt(row[0]), name: row[1], stock: parseFloat(row[3]) }));
  res.json(stock);
});

// Рецепты
app.get('/api/recipes', async (req, res) => {
  const rows = await getSheetData('РЕЦЕПТЫ!A2:D');
  const recipes = rows.map(row => ({ id: parseInt(row[0]), name: row[1], yield: parseFloat(row[2]), cost: parseFloat(row[3]) }));
  res.json(recipes);
});

// Финансы
app.get('/api/finance', async (req, res) => {
  const { startDate, endDate } = req.query;
  const rows = await getSheetData('ФИНАНСЫ!A2:F');
  let transactions = rows.map(row => ({
    id: parseInt(row[0]),
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
  res.json({ transactions, totalIncome, totalExpense, profit: totalIncome - totalExpense });
});

// Регистрация
app.post('/api/register', async (req, res) => {
  const { name, email, password } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: 'Заполните все поля' });
  const users = await getSheetData('ПОЛЬЗОВАТЕЛИ!A2:G');
  if (users.some(u => u[2] === email)) return res.status(400).json({ error: 'Пользователь уже существует' });
  const saltRounds = 10;
  const passwordHash = await bcrypt.hash(password, saltRounds);
  let lastId = 0;
  users.forEach(u => { if (u[0]) lastId = Math.max(lastId, parseInt(u[0])); });
  const newId = lastId + 1;
  const success = await appendRow('ПОЛЬЗОВАТЕЛИ!A:G', [newId, name, email, 'user', passwordHash, '', '']);
  if (success) res.json({ success: true, userId: newId });
  else res.status(500).json({ error: 'Ошибка регистрации' });
});
// Вход
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
// Выход
app.post('/api/logout', (req, res) => {
  res.clearCookie('access_token');
  res.json({ success: true });
});
// Текущий пользователь
app.get('/api/me', authenticateToken, (req, res) => {
  res.json(req.user);
});

app.listen(PORT, () => console.log(`🚀 Сервер запущен на порту ${PORT}`));
