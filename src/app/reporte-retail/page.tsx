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

// This is required for pdfjs-dist to work in the browser
pdfjs.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.min.mjs`;


export default function ReporteRetailPage() {
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<{pdf: string | null, excel: string | null}>({ pdf: null, excel: null });
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
  };

  const handleProcess = async () => {
    if (!pdfFile || !excelFile) {
      toast({ title: "Faltan archivos", description: "Por favor, sube ambos archivos, el PDF y el Excel.", variant: "destructive" });
      return;
    }

    setLoading(true);
    
    try {
        // 1. Parse Excel to create order -> document number map
        const excelBuffer = await excelFile.arrayBuffer();
        const workbook = XLSX.read(excelBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet) as any[];

        if (json.length === 0) {
            throw new Error("El archivo Excel está vacío o no tiene el formato esperado.");
        }
        
        const headers = Object.keys(json[0]);
        
        const orderHeader = headers.find(h => {
          const cleaned = h.trim().toLowerCase();
          return cleaned.includes('orden') || cleaned.includes('etiquetas de fila');
        });
        const docHeader = headers.find(h => {
          const cleaned = h.trim().toLowerCase();
          return cleaned.includes('documento') || cleaned.includes('ebeln');
        });

        if (!orderHeader || !docHeader) {
            throw new Error("El archivo Excel debe contener columnas para 'orden' ('Etiquetas de fila') y 'documento' ('EBELN').");
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
        const pdfBuffer = await pdfFile.arrayBuffer();
        const loadingTask = pdfjs.getDocument({ data: pdfBuffer });
        const pdfDocument = await loadingTask.promise;
        
        const pageInfo: { pageIndex: number; docNumber: number }[] = [];
        const unmappedPages: number[] = [];
        
        const pagePromises = [];
        for (let i = 0; i < pdfDocument.numPages; i++) {
            const pagePromise = pdfDocument.getPage(i + 1).then(async (page) => {
                const textContent = await page.getTextContent();
                const text = textContent.items.map((item: any) => item.str).join(' ');
                const orderRegex = /Orden[\s\S]*?(\d{10})/;
                const match = text.match(orderRegex);
                const orderNumber = match ? match[1] : null;
                return { pageIndex: i, orderNumber };
            }).catch(() => {
                return { pageIndex: i, orderNumber: null, error: true };
            });
            pagePromises.push(pagePromise);
        }
    
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
        
        // Sanity check to prevent out-of-bounds errors
        const validIndices = sortedPageIndices.filter(idx => idx < originalPdf.getPageCount());
        const copiedPages = await sortedPdf.copyPages(originalPdf, validIndices);
        copiedPages.forEach(page => sortedPdf.addPage(page));

        const sortedPdfBytes = await sortedPdf.save();
        const blob = new Blob([sortedPdfBytes], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);
        handleDownload(url, `ordenado_${pdfFile.name}`);

        toast({
          title: "Proceso completado",
          description: "El PDF ha sido reordenado. La descarga comenzará.",
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
          Sube un PDF de reportes y un archivo Excel con el mapeo de órdenes para generar un nuevo PDF con las páginas ordenadas por número de documento.
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
              Selecciona el archivo PDF que quieres ordenar y el archivo de Excel que contiene la relación entre "Orden" y "Número de Documento".
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">2</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">El Navegador Procesa</h4>
            <p className="text-foreground/80">
              Tu navegador lee el número de "Orden" de cada página del PDF, lo busca en tu Excel para encontrar el "Número de Documento" correspondiente.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">3</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Descarga el PDF Ordenado</h4>
            <p className="text-foreground/80">
              Se genera un nuevo archivo PDF con todas las páginas del original, pero reordenadas secuencialmente según el "Número de Documento".
            </p>
          </div>
        </div>
      </section>
    </main>
  );
}
