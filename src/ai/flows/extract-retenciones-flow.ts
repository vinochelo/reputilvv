'use server';
/**
 * @fileOverview Flow to extract tax retention data from a PDF.
 *
 * - extractRetenciones - A function that handles the PDF processing and data extraction.
 */

import { ai } from '@/ai/genkit';
import { z } from 'zod';
import pdf from 'pdf-parse';

const ExtractRetencionesInputSchema = z.object({
  pdfDataUri: z
    .string()
    .describe(
      "A PDF file of tax retentions, as a data URI that must include a MIME type and use Base64 encoding. Expected format: 'data:<mimetype>;base64,<encoded_data>'."
    ),
});

const extractorPrompt = ai.definePrompt({
  name: 'retencionesExtractorPrompt',
  input: { schema: z.object({ pdfText: z.string() }) },
  prompt: `You are an expert data extractor. Your task is to analyze the text from a tax withholding document from Ecuador and extract the required information in CSV format.

The columns must be exactly:
RUC,COMPROBANTE,COD. COMP.,Nro. Autorizacion,BASE IMPONIBLE,FECHA EMISION

Here are the details for each column:
- RUC: The issuer's RUC. Find the number associated with "RUC:".
- COMPROBANTE: The voucher number. It usually has a format like "001-001-000123456". You must extract only the full number without dashes (e.g., 001001000123456).
- COD. COMP.: The voucher type code. It is "01" for "FACTURA". Look for the document type.
- Nro. Autorizacion: The authorization number. It is a long numeric string.
- BASE IMPONIBLE: The taxable base amount for the withholding. Find the value associated with "Base Imponible para la Retención" or "IMPORTE BASE".
- FECHA EMISION: The issue date of the voucher. Find the date associated with "Fecha de Emisión" and format it as DD/MM/YYYY.

Example of input text:
---
RUC: 179...
FACTURA
001-102-000000123
...
FECHA DE EMISION: 15/01/2024
...
Nro. Autorizacion
150120240117...
...
Comprobante: FACTURA / 001-102-000000123
Impuesto a la Renta
Base Imponible para la Retención: 100.00
---

Expected CSV output for the example:
179...,001102000000123,01,150120240117...,100.00,15/01/2024

Analyze the following document text and provide ONLY the CSV output. Do not include a header row. If you find multiple withholdings, provide a new CSV line for each one.

Text from PDF:
{{{pdfText}}}`,
});

const extractRetencionesFlow = ai.defineFlow(
  {
    name: 'extractRetencionesFlow',
    inputSchema: ExtractRetencionesInputSchema,
    outputSchema: z.string(),
  },
  async ({ pdfDataUri }) => {
    // 1. Decode and parse the PDF
    const pdfBuffer = Buffer.from(pdfDataUri.split(',')[1], 'base64');
    const pdfData = await pdf(pdfBuffer);
    const pdfText = pdfData.text;

    if (!pdfText.trim()) {
      throw new Error("El PDF está vacío o no se pudo extraer texto.");
    }

    // 2. Call the AI to extract data as CSV
    const llmResponse = await extractorPrompt({ pdfText });
    const csvData = llmResponse.text;
    
    if (!csvData) {
      throw new Error("La IA no pudo extraer información del PDF.");
    }

    // 3. Process CSV to final TXT format
    const lines = csvData.trim().split('\n').filter(line => line.trim() !== '');

    const txtContent = lines.map(line => {
      const csvValues = line.split(',');
      if (csvValues.length < 6) return null;

      try {
        const ruc = csvValues[0].trim().padStart(13, '0');
        const comprobanteCompleto = csvValues[1].trim().padStart(15, '0');
        const tipoComp = csvValues[2].trim().padStart(2, '0');
        const serie = comprobanteCompleto.substring(0, 6);
        const comprobante = comprobanteCompleto.substring(6);
        const autorizacion = csvValues[3].trim();
        const baseImponible = parseFloat(csvValues[4].trim()).toFixed(2);
        const fechaParts = csvValues[5].trim().split('/');
        if (fechaParts.length < 3) return null;
        const fechaFormateada = `${fechaParts[0]}${fechaParts[1]}${fechaParts[2]}`;
        
        return `01|${fechaFormateada}|03|${ruc}|${tipoComp}|${serie}|${comprobante}|${autorizacion}|${baseImponible}|0|0|0`;
      } catch(e) {
        return null;
      }
    }).filter(line => line !== null).join('\r\n');

    if (!txtContent) {
      throw new Error("No se encontraron datos válidos para procesar después de la extracción. Revisa el formato del PDF.");
    }

    return txtContent;
  }
);

// Wrapper function to be called from the client
export async function extractRetenciones(pdfDataUri: string): Promise<string> {
    return extractRetencionesFlow({ pdfDataUri });
}
