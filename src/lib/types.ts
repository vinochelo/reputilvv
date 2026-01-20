export type ExcelData = {
  'Documento material': number;
  'Factura': string;
  'Nº doc.': number;
  'Centro': string;
  'Fecha Factura': string | Date;
  'Proveedor': number;
  'Nombre del proveedor': string;
  'Material': string;
  'Texto breve de material': string;
  'Cantidad': number;
  'Costo Total': number;
  'Precio Venta': number;
  'Utilidad %': number;
  'Valor a pagar': number;
  
  // From old ExcelData, for client info.
  'Rut': string;
  'Dv': string;
  'Nombre cliente': string;
};

export type GroupedData = {
  n_doc: number; // The group key
  items: ExcelData[];
  totalCantidad: number;
  totalCostoTotal: number;
  totalPrecioVenta: number;
  totalValorAPagar: number;
};
