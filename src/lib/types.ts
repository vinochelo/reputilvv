export type ExcelData = {
  'N doc': number;
  'Rut': string;
  'Dv': string;
  'Nombre cliente': string;
  'Tipo doc': string;
  'Folio': number;
  'Fecha emision': string;
  'Fecha venc': string;
  'Valor a pagar': number;
  'Vendedor': string;
};

export type GroupedData = {
  n_doc: number;
  cliente: string;
  rut: string;
  items: ExcelData[];
  totalPagar: number;
};
