import express, { Request, Response, Router } from 'express';
import ExcelJS from 'exceljs';
import { ExportExcelService } from '../exportExcelService/exportExcelService';

const router = Router();
const exportExcelService = new ExportExcelService()

router.get('/download-excel', async (req: Request, res: Response) => {
  try {
    await exportExcelService.handle(res)
    
  } catch (error) {
    console.error('Erro ao gerar o arquivo Excel:', error);
    res.status(500).send('Erro ao gerar o Excel');
  }
});

export default router
