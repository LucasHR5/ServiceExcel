import express, { Request, Response } from 'express';
import ExcelJS from 'exceljs';
import router from './routes/routes';

const app = express();
const PORT = 3000;

app.use('/Excel', router)
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
