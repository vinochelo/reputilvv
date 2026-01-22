'use server';
/**
 * @fileOverview Flow to reorder a PDF based on an Excel mapping.
 *
 * - sortRetailPdf - A function that handles the PDF and Excel processing and reordering.
 */

import { ai } from '@/ai/genkit';
import { z } from 'zod';
import * as XLSX from 'xlsx';
import { PDFDocument } from 'pdf-lib';
import pdf from 'pdf-parse';

// Define Zod schemas for input
const SortRetailPdfInputSchema = z.object({
  pdfDataUri: z
    .string()
    .describe("The retail report PDF file, as a data URI."),
  excelDataUri: z
    .string()
    .describe("The Excel mapping file, as a data URI."),
});


const sortRetailPdfFlow = ai.defineFlow(
  {
    name: 'sortRetailPdfFlow',
    inputSchema: SortRetailPdfInputSchema,
    outputSchema: z.string().describe("The reordered PDF as a data URI."),
  },
  async ({ pdfDataUri, excelDataUri }) => {
    // 1. Parse Excel to create order -> document number map
    const excelBuffer = Buffer.from(excelDataUri.split(',')[1], 'base64');
    const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet) as any[];

    if (json.length === 0) {
        throw new Error("El archivo Excel está vacío o no tiene el formato esperado.");
    }
    
    const headers = Object.keys(json[0]);
    const orderHeader = headers.find(h => h.toLowerCase().includes('orden'));
    const docHeader = headers.find(h => h.toLowerCase().includes('documento'));

    if (!orderHeader || !docHeader) {
        throw new Error("El archivo Excel debe contener columnas para 'orden' y 'documento'.");
    }

    const orderToDocMap = new Map<string, number>();
    for (const row of json) {
      const order = String(row[orderHeader]).trim();
      const docNum = Number(row[docHeader]);
      if (order && !isNaN(docNum)) {
        orderToDocMap.set(order, docNum);
      }
    }
    
    if (orderToDocMap.size === 0) {
        throw new Error("No se encontró un mapeo válido de orden a documento en el Excel.");
    }

    // 2. Parse PDF and extract order number from each page safely
    const pdfBuffer = Buffer.from(pdfDataUri.split(',')[1], 'base64');
    const pageInfo: { pageIndex: number; docNumber: number }[] = [];
    const unmappedPages: number[] = [];
    
    const pagePromises: Promise<{ pageIndex: number; orderNumber: string | null; error?: boolean }>[] = [];

    const render_page = (pageData: any) => {
        // Defensive check for invalid pageData object
        if (!pageData || typeof pageData.getTextContent !== 'function') {
            return ""; // Do nothing for invalid page data
        }

        const currentPageIndex = pageData.pageNumber - 1;

        const pagePromise = pageData.getTextContent().then((textContent: any) => {
            const text = (textContent?.items ?? []).map((item: any) => item.str).join(' ');
            const orderRegex = /Orden[\s\S]*?([0-9]{10})/;
            const match = text.match(orderRegex);
            const orderNumber = match ? match[1] : null;

            return { pageIndex: currentPageIndex, orderNumber };
        }).catch((err: any) => {
            // If text extraction fails for a page, return an error state but don't crash.
            return { pageIndex: currentPageIndex, orderNumber: null, error: true };
        });

        pagePromises.push(pagePromise);
        return ""; // Return value is not used by pdf-parse
    };
    
    await pdf(pdfBuffer, { pagerender: render_page });

    const extractedData = await Promise.all(pagePromises);

    for (const data of extractedData) {
        if (data.error) {
            unmappedPages.push(data.pageIndex);
            continue;
        }

        if (data.orderNumber && orderToDocMap.has(data.orderNumber)) {
            const docNumber = orderToDocMap.get(data.orderNumber)!;
            pageInfo.push({ pageIndex: data.pageIndex, docNumber });
        } else {
            unmappedPages.push(data.pageIndex);
        }
    }

    // 3. Sort pages based on document number
    pageInfo.sort((a, b) => a.docNumber - b.docNumber);
    
    // Combine sorted mapped pages with unmapped pages at the end
    const sortedPageIndices = [
        ...pageInfo.map(p => p.pageIndex),
        ...unmappedPages
    ];

    // 4. Create new PDF with sorted pages
    const originalPdf = await PDFDocument.load(pdfBuffer);
    const sortedPdf = await PDFDocument.create();
    const copiedPages = await sortedPdf.copyPages(originalPdf, sortedPageIndices);
    copiedPages.forEach(page => sortedPdf.addPage(page));

    const sortedPdfBytes = await sortedPdf.save();
    const sortedPdfDataUri = `data:application/pdf;base64,${Buffer.from(sortedPdfBytes).toString('base64')}`;

    return sortedPdfDataUri;
  }
);


// Wrapper function to be called from the client
export async function sortRetailPdf(pdfDataUri: string, excelDataUri: string): Promise<string> {
    return sortRetailPdfFlow({ pdfDataUri, excelDataUri });
}
