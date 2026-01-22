"use client";

import { useState, useRef, ChangeEvent } from 'react';
import { ArrowLeft, BrainCircuit, File as FileIcon, FileSpreadsheet, Loader2 } from 'lucide-react';
import Link from 'next/link';

import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { sortRetailPdf } from '@/ai/flows/sort-retail-pdf-flow';

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

  const fileToDataUri = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target?.result as string);
      reader.onerror = (e) => reject(e);
      reader.readAsDataURL(file);
    });
  }

  const handleProcess = async () => {
    if (!pdfFile || !excelFile) {
      toast({ title: "Faltan archivos", description: "Por favor, sube ambos archivos, el PDF y el Excel.", variant: "destructive" });
      return;
    }

    setLoading(true);
    
    try {
        const pdfDataUri = await fileToDataUri(pdfFile);
        const excelDataUri = await fileToDataUri(excelFile);

        const resultPdfDataUri = await sortRetailPdf(pdfDataUri, excelDataUri);

        toast({
          title: "Proceso completado",
          description: "El PDF ha sido reordenado. La descarga comenzará.",
        });

        handleDownload(resultPdfDataUri, `ordenado_${pdfFile.name}`);

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

  const handleDownload = (dataUri: string, name: string) => {
    const link = document.createElement('a');
    link.href = dataUri;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    // No es necesario revocar el object URL para data URIs
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
            <h4 className="text-xl font-semibold text-foreground">La IA Procesa la Información</h4>
            <p className="text-foreground/80">
              La herramienta lee el número de "Orden" de cada página del PDF, lo busca en tu Excel para encontrar el "Número de Documento" correspondiente.
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
