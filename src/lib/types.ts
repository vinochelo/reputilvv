export type ExcelData = {
  'Doc.material': number;
  'Factura': string;
  'Nº doc.': number;
  'Centro': string;
  'Fecha Factura': string;
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
  
  // The items in the group
  items: ExcelData[];

  // Some header info for the group
  cliente: string;
  rut: string;
  
  // Totals for the group
  totalCantidad: number;
  totalCostoTotal: number;
  totalPrecioVenta: number;
  totalValorAPagar: number;
};
