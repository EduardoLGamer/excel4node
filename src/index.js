const { Workbook } = require("excel4node");

const workbook = new Workbook();

const sheet = workbook.addWorksheet("Libro 1");

const titleStyle = workbook.createStyle({
  font: {
    color: "#000000",
    size: 13,
  },
});

const defaultStyle = workbook.createStyle({
  font: {
    color: "#000000",
    size: 12,
  },
  numberFormat: "$#,##0.00; ($#,##0.00); -",
});

const totalStyle = workbook.createStyle({
  font: {
    color: "#FF6C00",
    size: 12,
  },
  numberFormat: "$#,##0.00; ($#,##0.00); -",
});

sheet.cell(1, 1).string("Codigo").style(titleStyle);
sheet.cell(1, 2).string("Nombre").style(titleStyle);
sheet.cell(1, 3).string("Sub Nombre").style(titleStyle);
sheet.cell(1, 4).string("Costo de venta").style(titleStyle);
sheet.cell(1, 5).string("Cantidad").style(titleStyle);
sheet.cell(1, 6).string("Importe").style(titleStyle);

let items;
const items1 = [
  { code: "001", name: "Plumas", subName: "A", costSale: 10, quantity: 5 },
  { code: "002", name: "Colores", subName: "B", costSale: 15, quantity: 6 },
  { code: "003", name: "Lapiz ", subName: "---", costSale: 12, quantity: 3 },
];

const items2 = [
  { code: "001", name: "Plumas", subName: "A", costSale: 1, quantity: 7 },
  { code: "002", name: "Colores", subName: "B", costSale: 7, quantity: 2 },
  { code: "003", name: "Lapiz ", subName: "---", costSale: 5, quantity: 11 },
];

items = items1.concat(items2);

const endCell = items.length + 2;

for (let i = 0; i < items.length; i++) {
  const item = items[i];

  const cell = 2 + i;
  sheet.cell(cell, 1).string(item.code).style(defaultStyle);
  sheet.cell(cell, 2).string(item.name).style(defaultStyle);
  sheet.cell(cell, 3).string(item.subName).style(defaultStyle);
  sheet.cell(cell, 4).number(item.costSale).style(defaultStyle);
  sheet.cell(cell, 5).number(item.quantity).style(defaultStyle);

  sheet.cell(cell, 6).formula(`E${cell} * D${cell}`).style(defaultStyle);
}
let numbers = "";

for (let i = 0; i < endCell - 2; i++) {
  const cell = 2 + i;

  if (cell == endCell - 1) {
    numbers += `F${cell}`;
  } else {
    numbers += `F${cell} + `;
  }
}

sheet.cell(endCell, 6).formula(numbers).style(totalStyle);

workbook.write("new.xlsx");
