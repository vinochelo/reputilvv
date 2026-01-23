"use client";

import { useState, useRef, ChangeEvent } from 'react';
import { ArrowLeft, BrainCircuit, File as FileIcon, FileSpreadsheet, Loader2 } from 'lucide-react';
import Link from 'next/link';
import * as XLSX from 'xlsx';
import { PDFDocument } from 'pdf-lib';
import * as pdfjs from 'pdfjs-dist';

import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Separator } from '@/components/ui/separator';

// This is required for pdfjs-dist to work in the browser
pdfjs.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.min.mjs`;


export default function ReporteRetailPage() {
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<{pdf: string | null, excel: string | null}>({ pdf: null, excel: null });
  const [debugData, setDebugData] = useState<{rawText: string, orders: string[]} | null>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>, type: 'pdf' | 'excel') => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (type === 'pdf' && file.type !== 'application/pdf') {
      toast({ title: "Error de archivo", description: "Por favor, sube un archivo PDF.", variant: "destructive" });
      return;
    }

    if (type === 'excel' && !['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(file.type)) {
      toast({ title: "Error de archivo", description: "Por favor, sube un archivo de Excel (.xlsx o .xls).", variant: "destructive" });
      return;
    }
    
    if (type === 'pdf') setPdfFile(file);
    if (type === 'excel') setExcelFile(file);
    setFileName(prev => ({ ...prev, [type]: file.name }));
    setDebugData(null);
  };

  const handleProcess = async () => {
    if (!pdfFile || !excelFile) {
      toast({ title: "Faltan archivos", description: "Por favor, sube ambos archivos, el PDF y el Excel.", variant: "destructive" });
      return;
    }

    setLoading(true);
    setDebugData(null);
    
    try {
        // 1. Read and parse Excel
        const excelBuffer = await excelFile.arrayBuffer();
        const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        let headerRowIndex = -1;

        // Search for header row in the first 10 rows of the sheet
        for(let R = range.s.r; R <= Math.min(range.e.r, 10); ++R) {
            const rowValues = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell = worksheet[XLSX.utils.encode_cell({r:R, c:C})];
                if (cell && cell.v) {
                    rowValues.push(String(cell.v).trim().toLowerCase());
                }
            }
            // Check if both headers are present in the row
            if (rowValues.includes('ebeln') && rowValues.includes('belnr')) {
                headerRowIndex = R;
                break;
            }
        }

        if (headerRowIndex === -1) {
             throw new Error("No se pudo encontrar la fila de encabezado con 'EBELN' y 'BELNR' en el archivo Excel. Asegúrate de que el archivo exportado de SAP contenga estas columnas.");
        }
        
        // Parse the sheet into JSON starting from the found header row
        const json = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex }) as any[];

        if (json.length === 0) {
            throw new Error("El archivo Excel está vacío o no tiene el formato esperado después de encontrar los encabezados.");
        }
        
        // Find the exact header names (case-insensitive)
        const belnrHeader = Object.keys(json[0]).find(h => h.trim().toLowerCase() === 'belnr');
        const ebelnHeader = Object.keys(json[0]).find(h => h.trim().toLowerCase() === 'ebeln');


        if (!belnrHeader || !ebelnHeader) {
            const foundHeaders = Object.keys(json[0]).join(', ');
             setDebugData({
                rawText: `Columnas encontradas en el Excel: ${foundHeaders}`,
                orders: ["Se esperaban los encabezados: 'BELNR' y 'EBELN'"],
             })
            throw new Error(`El archivo Excel debe contener las columnas 'BELNR' y 'EBELN'.`);
        }

        // 2. Process Excel: deduplicate rows and sort by BELNR
        const uniquePairs = new Map<string, { belnr: string, ebeln: string }>();
        json.forEach((row: any) => {
            const belnr = String(row[belnrHeader]).trim();
            const ebeln = String(row[ebelnHeader]).trim();
            if (belnr && ebeln) {
                // Use a composite key to ensure uniqueness of the belnr-ebeln pair
                const key = `${belnr}-${ebeln}`;
                if (!uniquePairs.has(key)) {
                    uniquePairs.set(key, { belnr, ebeln });
                }
            }
        });

        let processedData = Array.from(uniquePairs.values());
        // Sort by BELNR, then EBELN as a secondary sort
        processedData.sort((a, b) => {
            const belnrCompare = a.belnr.localeCompare(b.belnr, undefined, { numeric: true });
            if (belnrCompare !== 0) return belnrCompare;
            return a.ebeln.localeCompare(b.ebeln, undefined, { numeric: true });
        });
        
        if (processedData.length === 0) {
            throw new Error("No se encontraron valores de 'EBELN' y 'BELNR' válidos para procesar en el Excel.");
        }
        
        // 3. Read PDF and map pages to EBELN values
        const pdfBufferForRead = await pdfFile.arrayBuffer();
        const loadingTask = pdfjs.getDocument({ data: pdfBufferForRead });
        const pdfDocument = await loadingTask.promise;
        
        const ebelnToPageIndices = new Map<string, number[]>();
        const allEbelnsFromExcel = [...new Set(processedData.map(p => p.ebeln))];

        const pageOwner: { [pageIndex: number]: string } = {};
        for (let i = 0; i < pdfDocument.numPages; i++) {
            const page = await pdfDocument.getPage(i + 1);
            const textContent = await page.getTextContent();
            const numericText = textContent.items.map((item: any) => item.str).join('').replace(/\D/g, '');

            for (const ebeln of allEbelnsFromExcel) {
                if (numericText.includes(ebeln)) {
                    pageOwner[i] = ebeln;
                    break; 
                }
            }
        }
        
        for (let i = 0; i < pdfDocument.numPages; i++) {
            const ownerEbeln = pageOwner[i];
            if (ownerEbeln) {
                if (!ebelnToPageIndices.has(ownerEbeln)) {
                    ebelnToPageIndices.set(ownerEbeln, []);
                }
                ebelnToPageIndices.get(ownerEbeln)!.push(i);
            }
        }

        if (Object.keys(pageOwner).length === 0 && pdfDocument.numPages > 0) {
             const firstPage = await pdfDocument.getPage(1);
             const textContent = await firstPage.getTextContent();
             const rawText = textContent.items.map((item: any) => item.str).join(' ');
             setDebugData({
                 rawText: rawText,
                 orders: allEbelnsFromExcel,
             });
        }

        // 4. Create new PDF with sorted pages
        const finalPageIndices: number[] = [];
        const usedEbelns = new Set<string>();
        
        for (const data of processedData) {
            const ebeln = data.ebeln;
            
            if (usedEbelns.has(ebeln)) {
                continue;
            }

            const pageIndicesForEbeln = ebelnToPageIndices.get(ebeln);
            
            if (pageIndicesForEbeln && pageIndicesForEbeln.length > 0) {
                finalPageIndices.push(...pageIndicesForEbeln);
                usedEbelns.add(ebeln);
            }
        }
        
        const allMappedIndices = new Set(finalPageIndices);
        const unmappedPagesCount = pdfDocument.numPages - allMappedIndices.size;
        
        for (let i = 0; i < pdfDocument.numPages; i++) {
            if (!allMappedIndices.has(i)) {
                finalPageIndices.push(i);
            }
        }
        
        const pdfBufferForWrite = await pdfFile.arrayBuffer();
        const originalPdf = await PDFDocument.load(pdfBufferForWrite);
        const sortedPdf = await PDFDocument.create();
        
        const copiedPages = await sortedPdf.copyPages(originalPdf, finalPageIndices);
        copiedPages.forEach((page) => {
            sortedPdf.addPage(page);
        });

        const sortedPdfBytes = await sortedPdf.save();
        const blob = new Blob([sortedPdfBytes], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);
        handleDownload(url, `ordenado_${pdfFile.name}`);

        toast({
          title: "Proceso completado",
          description: `Se reordenaron ${allMappedIndices.size} de ${pdfDocument.numPages} páginas. ${unmappedPagesCount} páginas no se pudieron mapear y se añadieron al final.`,
        });

    } catch (error: any) {
      console.error("Error processing files:", error);
      toast({
        title: "Error al procesar los archivos",
        description: error.message || "No se pudo completar el proceso. Revisa los archivos e intenta de nuevo.",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = (url: string, name: string) => {
    const link = document.createElement('a');
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };


  return (
     <main className="container mx-auto px-4 py-12">
      <div className="mb-8">
        <Link href="/" className="inline-flex items-center text-sm font-medium text-primary hover:underline">
          <ArrowLeft className="mr-2 h-4 w-4" />
          Volver al portal
        </Link>
      </div>

      <div className="text-center mb-12">
        <h1 className="text-4xl font-headline font-bold tracking-tight text-foreground sm:text-5xl lg:text-6xl">
          Reportes de Retail
        </h1>
        <p className="mt-4 max-w-3xl mx-auto text-lg text-foreground/80">
          Sube un PDF y un Excel. La herramienta procesará el Excel, lo ordenará y luego reordenará tu PDF para que coincida con ese orden.
        </p>
      </div>

      <Card className="max-w-4xl mx-auto shadow-lg">
        <CardHeader>
          <CardTitle>Organizador de Reportes Retail</CardTitle>
          <CardDescription>Sube los dos archivos requeridos para comenzar el proceso.</CardDescription>
        </CardHeader>
        <CardContent className="p-6 grid md:grid-cols-2 gap-8">
          <div className="flex flex-col items-center justify-center space-y-4 p-6 border-2 border-dashed rounded-lg">
            <FileIcon className="w-12 h-12 text-primary" />
            <p className="text-lg font-semibold text-foreground">1. Sube el reporte en PDF</p>
            <Button variant="outline" onClick={() => pdfInputRef.current?.click()}>
              Seleccionar PDF
            </Button>
            <input ref={pdfInputRef} type="file" className="hidden" onChange={(e) => handleFileChange(e, 'pdf')} accept="application/pdf" />
            {fileName.pdf && <p className="text-sm text-muted-foreground">{fileName.pdf}</p>}
          </div>

          <div className="flex flex-col items-center justify-center space-y-4 p-6 border-2 border-dashed rounded-lg">
            <FileSpreadsheet className="w-12 h-12 text-primary" />
            <p className="text-lg font-semibold text-foreground">2. Sube el reporte de Excel</p>
            <Button variant="outline" onClick={() => excelInputRef.current?.click()}>
              Seleccionar Excel
            </Button>
            <input ref={excelInputRef} type="file" className="hidden" onChange={(e) => handleFileChange(e, 'excel')} accept=".xlsx, .xls" />
            {fileName.excel && <p className="text-sm text-muted-foreground">{fileName.excel}</p>}
          </div>
        </CardContent>
        <div className="p-6 pt-0 flex justify-center">
            {loading ? (
                 <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <Loader2 className="w-12 h-12 text-primary animate-spin" />
                    <p className="text-lg font-semibold text-foreground">Procesando y reordenando el PDF...</p>
                    <p className="text-sm text-muted-foreground">Esto puede tomar un momento.</p>
                </div>
            ) : (
                <Button size="lg" onClick={handleProcess} disabled={!pdfFile || !excelFile}>
                  <BrainCircuit className="mr-2 h-5 w-5" />
                  Ordenar PDF
                </Button>
            )}
        </div>
      </Card>

      {debugData && (
        <Card className="max-w-4xl mx-auto mt-8">
          <CardHeader>
            <CardTitle>Información de Depuración</CardTitle>
            <CardDescription>
              No se pudo ordenar ninguna página. Revisa que el 'EBELN' del Excel exista en el PDF.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div>
              <h4 className="font-semibold mb-2">Texto Extraído de la Primera Página del PDF:</h4>
              <pre className="p-4 bg-muted rounded-md text-xs whitespace-pre-wrap max-h-[300px] overflow-auto">
                {debugData.rawText}
              </pre>
            </div>
            <div>
              <h4 className="font-semibold mb-2">Valores 'EBELN' buscados (del archivo Excel):</h4>
              <p className="p-4 bg-muted rounded-md text-xs">
                {debugData.orders.join(', ')}
              </p>
            </div>
          </CardContent>
        </Card>
      )}

      <section className="max-w-4xl mx-auto mt-12">
        <Accordion type="single" collapsible className="w-full">
          <AccordionItem value="item-1">
            <AccordionTrigger className="text-xl font-headline font-semibold text-primary hover:no-underline">
              Instrucciones para obtener los archivos de SAP
            </AccordionTrigger>
            <AccordionContent>
              <div className="space-y-6 pt-4 text-base">
                <div>
                  <h4 className="font-bold text-lg mb-3 text-foreground/90">Parte 1: Generar el Reporte en PDF</h4>
                  <ol className="list-decimal list-inside space-y-4 text-foreground/80">
                    <li>
                      <strong>Transacción `MIR5`</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Ingresa, filtra y descarga las facturas según el rango de fechas, usuario y sociedad.</li>
                        <li>Copia o exporta los <strong>números de documento (`BELNR`)</strong>, los necesitarás en el siguiente paso.</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Transacción `ZREP PEDIDOS`</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>En el campo de selección de "Facturas", pega todos los números de documento que obtuviste.</li>
                        <li>Ejecuta el reporte (F8).</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Generar el PDF</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Haz clic en el icono de <strong>REPORTE</strong> (o presiona `Shift+F1`).</li>
                        <li>En la ventana de impresión, elige la impresora <strong>"Microsoft Print to PDF"</strong>.</li>
                        <li>Imprime y guarda el archivo. Este será el PDF que subirás a la herramienta.</li>
                      </ul>
                    </li>
                  </ol>
                </div>
                <Separator />
                <div>
                  <h4 className="font-bold text-lg mb-3 text-foreground/90">Parte 2: Generar el Archivo Excel</h4>
                  <ol className="list-decimal list-inside space-y-4 text-foreground/80">
                    <li>
                      <strong>Transacción `SE16`</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Ingresa a la transacción `SE16`, escribe la tabla <strong>`EKBE`</strong> y presiona Enter.</li>
                        <li>Carga la variante: Menú <strong>Pasar a &gt; Variantes &gt; Traer...</strong> y selecciona <strong>`REVOC`</strong>.</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Filtrar Documentos</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>En el campo <strong>`BELNR`</strong>, usa la selección múltiple para pegar todos los números de documento de la `MIR5`.</li>
                        <li>Ejecuta la selección (F8).</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Exportar a Excel</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Asegúrate de que las columnas `BELNR` y `EBELN` estén visibles.</li>
                        <li>Ve a <strong>Sistema &gt; Lista &gt; Grabar &gt; Fichero local</strong>.</li>
                        <li>Elige la opción <strong>"Hoja de cálculo"</strong>.</li>
                        <li>Guarda el archivo. Este será el archivo Excel que subirás a la herramienta. <strong>No necesitas crear una tabla dinámica.</strong></li>
                      </ul>
                    </li>
                  </ol>
                </div>
              </div>
            </AccordionContent>
          </AccordionItem>
        </Accordion>
      </section>

       <section className="mt-16 max-w-4xl mx-auto">
        <h3 className="text-3xl font-headline font-bold text-center mb-8 text-foreground">
          ¿Cómo funciona?
        </h3>
        <div className="grid md:grid-cols-3 gap-8 text-center">
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">1</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Sube tus Archivos</h4>
            <p className="text-foreground/80">
              Selecciona el PDF y el Excel extraído de SAP. El Excel debe tener las columnas 'BELNR' y 'EBELN'.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">2</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">La Herramienta Procesa</h4>
            <p className="text-foreground/80">
              La aplicación lee tu Excel, elimina duplicados y lo ordena. Luego, busca cada 'EBELN' en tu PDF para saber cómo reordenar las páginas.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">3</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Descarga el PDF Ordenado</h4>
            <p className="text-foreground/80">
              Se genera un nuevo archivo PDF con todas las páginas reordenadas secuencialmente según el orden de tu Excel.
            </p>
          </div>
        </div>
      </section>
    </main>
  );
}
