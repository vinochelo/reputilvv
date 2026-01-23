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
        const json = XLSX.utils.sheet_to_json(worksheet) as any[];

        if (json.length === 0) {
            throw new Error("El archivo Excel está vacío o no tiene el formato esperado.");
        }
        
        const sortHeader = Object.keys(json[0]).find(h => h.trim().toLowerCase() === 'etiquetas de fila');
        const searchHeader = Object.keys(json[0]).find(h => h.trim().toLowerCase() === 'ebeln');


        if (!sortHeader || !searchHeader) {
            const foundHeaders = Object.keys(json[0]).join(', ');
             setDebugData({
                rawText: `Columnas encontradas en el Excel: ${foundHeaders}`,
                orders: ["Se esperaban los encabezados: 'Etiquetas de fila' y 'EBELN'"],
             })
            throw new Error(`El archivo Excel debe contener las columnas 'Etiquetas de fila' (para ordenar) y 'EBELN' (para buscar).`);
        }

        const searchToSortMap = new Map<string, number>();
        json.forEach((row: any) => {
            const searchValue = String(row[searchHeader]).trim();
            const sortValue = Number(String(row[sortHeader]).trim());
            if (searchValue && !isNaN(sortValue)) {
                searchToSortMap.set(searchValue, sortValue);
            }
        });
        
        if (searchToSortMap.size === 0) {
            throw new Error("No se encontró un mapeo válido de 'EBELN' a 'Etiquetas de fila' en el Excel.");
        }
        
        // 2. Read PDF and find search values (EBELN)
        const pdfBuffer = await pdfFile.arrayBuffer();
        const loadingTask = pdfjs.getDocument({ data: pdfBuffer });
        const pdfDocument = await loadingTask.promise;
        
        const pageInfo: { pageIndex: number; sortValue: number }[] = [];
        const unmappedPages: number[] = [];
        
        const excelSearchValues = Array.from(searchToSortMap.keys());

        for (let i = 0; i < pdfDocument.numPages; i++) {
            const page = await pdfDocument.getPage(i + 1);
            const textContent = await page.getTextContent();
            
            const pageText = textContent.items.map((item: any) => item.str).join('');

            let foundSearchValue: string | null = null;
            
            const foundExcelValue = excelSearchValues.find(excelValue => pageText.includes(excelValue));

            if (foundExcelValue) {
                 foundSearchValue = foundExcelValue;
            }
            
            if (foundSearchValue && searchToSortMap.has(foundSearchValue)) {
                const sortValue = searchToSortMap.get(foundSearchValue)!;
                pageInfo.push({ pageIndex: i, sortValue });
            } else {
                unmappedPages.push(i);
            }
        }

        if (pageInfo.length === 0 && pdfDocument.numPages > 0) {
             const firstPage = await pdfDocument.getPage(1);
             const textContent = await firstPage.getTextContent();
             const rawText = textContent.items.map((item: any) => item.str).join(' ');
             setDebugData({
                 rawText: rawText,
                 orders: excelSearchValues, // These are the EBELN values
             });
            toast({
                title: "No se pudo ordenar",
                description: `Se ordenaron 0 de ${pdfDocument.numPages} páginas. Revisa la información de depuración.`,
                variant: "destructive",
            });
            setLoading(false);
            return;
        }

        // 3. Sort pages based on sort value (from 'Etiquetas de fila')
        pageInfo.sort((a, b) => a.sortValue - b.sortValue);
        
        const sortedPageIndices = [
            ...pageInfo.map(p => p.pageIndex),
            ...unmappedPages
        ];
        
        // 4. Create new PDF with sorted pages
        const pdfBufferForPdfLib = await pdfFile.arrayBuffer(); // Re-read buffer
        const originalPdf = await PDFDocument.load(pdfBufferForPdfLib);
        const sortedPdf = await PDFDocument.create();
        
        for (const pageIndex of sortedPageIndices) {
            const [copiedPage] = await sortedPdf.copyPages(originalPdf, [pageIndex]);
            sortedPdf.addPage(copiedPage);
        }

        const sortedPdfBytes = await sortedPdf.save();
        const blob = new Blob([sortedPdfBytes], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);
        handleDownload(url, `ordenado_${pdfFile.name}`);

        toast({
          title: "Proceso completado",
          description: `Se reordenaron ${pageInfo.length} de ${pdfDocument.numPages} páginas. ${unmappedPages.length} páginas no se pudieron mapear y se añadieron al final.`,
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
          Sube un PDF y un Excel. La herramienta buscará el valor 'EBELN' del Excel en cada página del PDF y luego ordenará las páginas según el valor 'Etiquetas de fila'.
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
            <p className="text-lg font-semibold text-foreground">2. Sube el mapeo en Excel</p>
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
              No se pudo ordenar ninguna página o hubo un error. Aquí está la información relevante para el diagnóstico.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div>
              <h4 className="font-semibold mb-2">Texto Extraído de la Primera Página del PDF (o de columnas del Excel):</h4>
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
                      <strong>Seleccionar Campos</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Haz clic derecho sobre los datos, elige <strong>"Selección de campo"</strong>, desmarca todo y luego selecciona solo <strong>`EBELN`</strong>, <strong>`GJAHR`</strong> y <strong>`BELNR`</strong>.</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Exportar a HTML/XLS</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Ve a <strong>Sistema &gt; Lista &gt; Grabar &gt; Grabar</strong>.</li>
                        <li>Elige <strong>"Formato HTML"</strong> y continúa.</li>
                        <li>Nombra el fichero asegurándote de que termine en <strong>`.xls`</strong> y guárdalo.</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Crear Tabla Dinámica en Excel</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Abre el archivo `.xls` (ignora las advertencias de formato) y guárdalo como <strong>Libro de Excel (.xlsx)</strong>.</li>
                        <li>Inserta una <strong>Tabla Dinámica</strong> con los datos.</li>
                        <li>Arrastra a <strong>Filas</strong>: primero <strong>`BELNR`</strong>, y debajo <strong>`EBELN`</strong>.</li>
                        <li>En la pestaña <strong>Diseño</strong> (de la tabla dinámica), ve a <strong>Diseño de Informe</strong> y elige <strong>"Mostrar en formato Tabular"</strong>.</li>
                        <li>En la misma pestaña, ve a <strong>Subtotales</strong> y elige <strong>"No mostrar subtotales"</strong>.</li>
                        <li>Tu tabla final debe tener dos columnas: "Etiquetas de fila" (que es `BELNR`) y "EBELN". Guarda este archivo. Este será el Excel que subirás.</li>
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
              Selecciona el PDF y el Excel. El Excel debe tener las columnas 'Etiquetas de fila' y 'EBELN'.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">2</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">El Navegador Procesa</h4>
            <p className="text-foreground/80">
              Tu navegador lee el número 'EBELN' de cada página del PDF, lo busca en tu Excel y usa 'Etiquetas de fila' para saber cómo ordenar.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">3</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Descarga el PDF Ordenado</h4>
            <p className="text-foreground/80">
              Se genera un nuevo archivo PDF con todas las páginas reordenadas secuencialmente según el valor de 'Etiquetas de fila'.
            </p>
          </div>
        </div>
      </section>
    </main>
  );
}
