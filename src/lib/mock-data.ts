import type { ExcelData } from '@/lib/types';

export const mockExcelData: ExcelData[] = [
  { 'N doc': 12345, 'Rut': '11.111.111', 'Dv': '1', 'Nombre cliente': 'Cliente Uno', 'Tipo doc': 'FC', 'Folio': 101, 'Fecha emision': '2023-10-01', 'Fecha venc': '2023-10-31', 'Valor a pagar': 50000, 'Vendedor': 'Vendedor A' },
  { 'N doc': 12345, 'Rut': '11.111.111', 'Dv': '1', 'Nombre cliente': 'Cliente Uno', 'Tipo doc': 'FC', 'Folio': 102, 'Fecha emision': '2023-10-02', 'Fecha venc': '2023-11-01', 'Valor a pagar': 75000, 'Vendedor': 'Vendedor A' },
  { 'N doc': 12346, 'Rut': '22.222.222', 'Dv': '2', 'Nombre cliente': 'Cliente Dos', 'Tipo doc': 'FC', 'Folio': 201, 'Fecha emision': '2023-10-05', 'Fecha venc': '2023-11-04', 'Valor a pagar': 120000, 'Vendedor': 'Vendedor B' },
  { 'N doc': 12346, 'Rut': '22.222.222', 'Dv': '2', 'Nombre cliente': 'Cliente Dos', 'Tipo doc': 'FC', 'Folio': 202, 'Fecha emision': '2023-10-06', 'Fecha venc': '2023-11-05', 'Valor a pagar': 30000, 'Vendedor': 'Vendedor B' },
  { 'N doc': 12346, 'Rut': '22.222.222', 'Dv': '2', 'Nombre cliente': 'Cliente Dos', 'Tipo doc': 'BL', 'Folio': 203, 'Fecha emision': '2023-10-07', 'Fecha venc': '2023-11-06', 'Valor a pagar': 15000, 'Vendedor': 'Vendedor B' },
  { 'N doc': 12347, 'Rut': '33.333.333', 'Dv': '3', 'Nombre cliente': 'Cliente Tres', 'Tipo doc': 'FC', 'Folio': 301, 'Fecha emision': '2023-10-10', 'Fecha venc': '2023-11-09', 'Valor a pagar': 200000, 'Vendedor': 'Vendedor C' },
  { 'N doc': 12347, 'Rut': '33.333.333', 'Dv': '3', 'Nombre cliente': 'Cliente Tres', 'Tipo doc': 'NC', 'Folio': 302, 'Fecha emision': '2023-10-11', 'Fecha venc': '2023-11-10', 'Valor a pagar': -50000, 'Vendedor': 'Vendedor C' },
];
