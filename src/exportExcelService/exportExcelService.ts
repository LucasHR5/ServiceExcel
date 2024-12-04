import ExcelJS from "exceljs";
import { Response } from "express";

const users = [
  {
    name: "Lucas Rodrigues",
    age: 27,
    date: new Date("1996-05-12"),
    phone: "+55 (11) 91234-5678",
    password: "senhaSegura123@",
  },
  {
    name: "Maria Silva",
    age: 30,
    date: new Date("1993-03-25"),
    phone: "+55 (21) 98765-4321",
    password: "minhaSenha123#",
  },
  {
    name: "João Santos",
    age: 22,
    date: new Date("2001-10-05"),
    phone: "+55 (31) 99876-5432",
    password: "123senhaJoao!",
  },
  {
    name: "Ana Oliveira",
    age: 35,
    date: new Date("1988-08-19"),
    phone: "+55 (51) 91123-4567",
    password: "anaSenhaForte$",
  },
  {
    name: "Carlos Pereira",
    age: 28,
    date: new Date("1995-12-01"),
    phone: "+55 (41) 91987-6543",
    password: "carlosSenha@@@",
  },
];
export class ExportExcelService {
  async handle(res: Response) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(" Users ");

    // Configurar as colunas e adicionar dados
    const row = worksheet.columns = [
      { header: "Nome", key: "name", width: 15, style: { alignment: { horizontal: "center"} } },
      { header: "Idade", key: "age", width: 10,  style: { alignment: { horizontal: "center"} } },
      { header: "data de nascimento", key: "date", width: 20,  style: { alignment: { horizontal: "center"} } },
      { header: "telefone", key: "phone", width: 20,  style: { alignment: { horizontal: "center"} } },
      { header: "Senha", key: "password", width: 20,  style: { alignment: { horizontal: "center"} } }, 
    ];

    

    

    users.forEach((user, index)=>{
      const row = worksheet.addRow({
        name: user.name,
        age: user.age,
        date: user.date.toLocaleDateString(),
        phone: user.phone,
        password: user.password
      });

      const fillColor = index %2 ===0 ? "FFFFFF" : "F2F2F2";
      
      row.eachCell((cell)=> {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: fillColor}
        } as ExcelJS.Fill
        cell.alignment = { vertical: "middle", horizontal: "center"};
        worksheet.getRow(1).font = { bold: true }
      })
    });

   

    // Configurar cabeçalhos para forçar o download
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=relatorio.xlsx");

    await workbook.xlsx.write(res);
    res.end();
  }
}
