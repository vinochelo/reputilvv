
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const mockExcelData = [
  {
    'Doc.material': 4906828376,
    'Factura': '001001000164749',
    'Nº doc.': 5200028672,
    'Centro': 'CD02',
    'Fecha Factura': '2026-01-12',
    'Proveedor': 110001744,
    'Nombre del proveedor': 'NEXSYS DEL ECUADOR',
    'Material': '5000001087708001',
    'Texto breve de material': 'LAP VIVOR7-5825U512-SILVER, TALLA UNIC',
    'Cantidad': 1,
    'Costo Total': 385,
    'Precio Venta': 597.35,
    'Utilidad %': 55.00,
    'Valor a pagar': 436.97,
  },
  {
    'Doc.material': 4906828377,
    'Factura': '001001000164750',
    'Nº doc.': 5200028672,
    'Centro': 'CD02',
    'Fecha Factura': '2026-01-12',
    'Proveedor': 110001744,
    'Nombre del proveedor': 'NEXSYS DEL ECUADOR',
    'Material': '5000001087708002',
    'Texto breve de material': 'MOUSE LOGITECH G502',
    'Cantidad': 2,
    'Costo Total': 100,
    'Precio Venta': 150,
    'Utilidad %': 50.00,
    'Valor a pagar': 120,
  },
  {
    'Doc.material': 4906828378,
    'Factura': '001001000164751',
    'Nº doc.': 5200028673,
    'Centro': 'CD03',
    'Fecha Factura': '2026-01-13',
    'Proveedor': 110001745,
    'Nombre del proveedor': 'PROVEEDOR OTRO',
    'Material': '5000001087708003',
    'Texto breve de material': 'TECLADO MECANICO',
    'Cantidad': 5,
    'Costo Total': 500,
    'Precio Venta': 750,
    'Utilidad %': 50.00,
    'Valor a pagar': 600,
  },
];

const worksheet = XLSX.utils.json_to_sheet(mockExcelData);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Resultados');

const filePath = path.join(__dirname, 'test_sap_data.xlsx');
XLSX.writeFile(workbook, filePath);

console.log(`Fichero de prueba creado en: ${filePath}`);
