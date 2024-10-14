const xl = require("excel4node");
const wb = new xl.Workbook();
const ws = wb.addWorksheet("Nome da Planilha");

const data = [
  {
    name: "John",
    age: 30,
    email: "john@example.com",
  },
  {
    name: "Jane",
    age: "25",
    email: "ana@example.com",
  },
];

const headingColumnNames = ["Name", "Age", "E-mail"];

// Insere nome nas colunas
headingColumnNames.forEach((heading, index) => {
  ws.cell(1, index + 1).string(heading);
});

// Insere os dados
data.forEach((row, index) => {
  Object.keys(row).forEach((key, colIndex) => {
    ws.cell(index + 2, colIndex + 1).string(String(row[key]));
  });
});

wb.write("planilha.xlsx");
