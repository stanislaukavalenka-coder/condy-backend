require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

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

// Вспомогательная функция для безопасного чтения диапазона
async function getSheetData(range) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
    });
    return response.data.values || [];
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

// --- Публичные маршруты ---
app.get('/ping', (req, res) => {
  res.json({ message: 'Сервер работает' });
});

// Заказы (чтение)
app.get('/api/orders', async (req, res) => {
  const rows = await getSheetData('ЗАКАЗЫ!A2:H');
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

// Создание заказа (без авторизации)
app.post('/api/orders', async (req, res) => {
  const { clientId, price, status, details, delivery, executionDate } = req.body;
  if (!clientId || !price || !executionDate) {
    return res.status(400).json({ error: 'Не хватает полей' });
  }
  // Определить следующий ID
  const rows = await getSheetData('ЗАКАЗЫ!A:A');
  let lastId = 0;
  rows.forEach(row => { const id = parseInt(row[0]); if (id > lastId) lastId = id; });
  const newId = lastId + 1;
  const now = new Date().toISOString();
  const success = await appendRow('ЗАКАЗЫ!A:H', [
    newId, now, clientId, price, status || 'в работе', details || '', delivery || '', executionDate
  ]);
  if (success) res.json({ success: true, orderId: newId });
  else res.status(500).json({ error: 'Ошибка создания заказа' });
});

// Клиенты
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

// Склад
app.get('/api/stock', async (req, res) => {
  const rows = await getSheetData('СКЛАД!A2:D');
  const stock = rows.map(row => ({
    id: row[0],
    name: row[1],
    stock: row[3],
  }));
  res.json(stock);
});

// Рецепты
app.get('/api/recipes', async (req, res) => {
  const rows = await getSheetData('РЕЦЕПТЫ!A2:D');
  const recipes = rows.map(row => ({
    id: row[0],
    name: row[1],
    yield: row[2],
    cost: row[3],
  }));
  res.json(recipes);
});

// Финансы
app.get('/api/finance', async (req, res) => {
  const { startDate, endDate } = req.query;
  const rows = await getSheetData('ФИНАНСЫ!A2:F');
  const transactions = rows.map(row => ({
    id: row[0],
    date: row[1],
    type: row[2],
    category: row[3],
    amount: parseFloat(row[4]),
    comment: row[5],
  }));
  let filtered = transactions;
  if (startDate && endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    filtered = transactions.filter(t => {
      const d = new Date(t.date);
      return d >= start && d <= end;
    });
  }
  const totalIncome = filtered.filter(t => t.type === 'Доход').reduce((s, t) => s + t.amount, 0);
  const totalExpense = filtered.filter(t => t.type === 'Расход').reduce((s, t) => s + t.amount, 0);
  res.json({
    transactions: filtered,
    totalIncome,
    totalExpense,
    profit: totalIncome - totalExpense,
  });
});

app.listen(PORT, () => {
  console.log(`🚀 Сервер запущен на порту ${PORT}`);
});
