require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;

// Настройка авторизации для Google Sheets
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// Проверочный маршрут
app.get('/ping', (req, res) => {
  res.json({ message: 'Сервер работает' });
});

// Получить все заказы из таблицы
app.get('/api/orders', async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'ЗАКАЗЫ!A2:I',
    });
    const rows = response.data.values || [];
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
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Ошибка чтения таблицы' });
  }
});

app.listen(PORT, () => {
  console.log(`Сервер запущен на порту ${PORT}`);
});