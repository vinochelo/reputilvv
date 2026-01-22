
"use client";

import { useState, useRef, ChangeEvent } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, FileX2, ArrowLeft } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow, TableFooter } from '@/components/ui/table';
import { Progress } from '@/components/ui/progress';
import { useToast } from '@/hooks/use-toast';
import type { ExcelData, GroupedData } from '@/lib/types';
import Link from 'next/link';

export default function ReporteVentaVerdePage() {
  const [processedData, setProcessedData] = useState<GroupedData[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [progress, setProgress] = useState(0);
  const pdfContentRef = useRef<HTMLDivElement>(null);
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const getDocId = (item: any): number | string | undefined => {
    const keys = Object.keys(item);
    const docIdRegex = /n(º|°|ro\.?|umero)?\.?\s*doc/i;
    for (const key of keys) {
        if (docIdRegex.test(key.trim())) {
            const value = item[key];
            if (value !== null && value !== undefined) {
                return value;
            }
        }
    }
    return undefined;
  };
  
  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      if (file.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' && file.type !== 'application/vnd.ms-excel') {
        toast({
          title: "Error de archivo",
          description: "Por favor, sube un archivo de Excel (.xlsx o .xls).",
          variant: "destructive",
        });
        return;
      }
      setFileName(file.name);
      setParsing(true);

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json<ExcelData>(worksheet);
          
          processData(jsonData);
        } catch (error) {
          console.error("Error parsing Excel file:", error);
          toast({
            title: "Error al procesar el archivo",
            description: "No se pudo leer el archivo de Excel. Asegúrate de que el formato sea correcto.",
            variant: "destructive",
          });
          resetState();
        } finally {
          setParsing(false);
        }
      };
      reader.onerror = () => {
        console.error("Error reading file");
        toast({
            title: "Error al leer el archivo",
            description: "Ocurrió un error al intentar leer el archivo.",
            variant: "destructive",
        });
        setParsing(false);
        resetState();
      }
      reader.readAsArrayBuffer(file);
    }
  };

  const processData = (data: ExcelData[]) => {
    const grouped = data.reduce<GroupedData[]>((acc, item) => {
      const docId = getDocId(item);

      if (docId === undefined || docId === null) {
        return acc; // Skip items without a document number.
      }

      let group = acc.find(g => g.n_doc === docId);

      if (!group) {
        group = {
          n_doc: docId,
          items: [],
          totalCantidad: 0,
          totalCostoTotal: 0,
          totalPrecioVenta: 0,
          totalValorAPagar: 0,
          totalUtilidad: 0,
        };
        acc.push(group);
      }
      
      const cantidad = Number(item['Cantidad'] ?? 0);
      const costoTotal = Number(item['Costo Total'] ?? 0);
      const precioVenta = Number(item['Precio Venta'] ?? 0);
      const utilidad = Number(item['Utilidad %'] ?? 0);
      const valorAPagar = Number(item['Valor a pagar'] ?? 0);
      
      group.items.push({
          ...item,
          'Cantidad': cantidad,
          'Costo Total': costoTotal,
          'Precio Venta': precioVenta,
          'Utilidad %': utilidad,
          'Valor a pagar': valorAPagar,
      });

      group.totalCantidad += cantidad;
      group.totalCostoTotal += costoTotal;
      group.totalPrecioVenta += precioVenta;
      group.totalUtilidad += utilidad;
      group.totalValorAPagar += valorAPagar;

      return acc;
    }, []);

    if (grouped.length === 0 && data.length > 0) {
      const firstRowKeys = Object.keys(data[0]).join(', ');
      toast({
        title: "No se pudo agrupar",
        description: `No se encontraron datos para agrupar. Revisa que la columna 'Nº doc.' exista y tenga valores. Columnas encontradas: ${firstRowKeys}`,
        variant: "destructive",
      });
      setProcessedData(null);
      return;
    }

    setProcessedData(grouped);
  };

  const handleDownloadPdf = async () => {
    if (!pdfContentRef.current) return;
  
    setLoading(true);
    setProgress(0);
    await new Promise(resolve => setTimeout(resolve, 50));
    
    const pdf = new jsPDF('l', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const reportElements = Array.from(pdfContentRef.current.children);
    let successfulPages = 0;
    
    for (let i = 0; i < reportElements.length; i++) {
      const reportElement = reportElements[i] as HTMLElement;
      if (reportElement) {
        try {
            const canvas = await html2canvas(reportElement, { scale: 2, useCORS: true });
            const imgData = canvas.toDataURL('image/png', 1.0); 
            
            if (i > 0) {
              pdf.addPage();
            }
  
            const imgProps = pdf.getImageProperties(imgData);
            const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
            pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);
            successfulPages++;
            setProgress(((i + 1) / reportElements.length) * 100);
            
            await new Promise(resolve => setTimeout(resolve, 50));
        } catch(e) {
            console.error("Error processing page for PDF:", e);
            toast({
              title: "Error al generar página",
              description: `No se pudo procesar el reporte ${i + 1} para el PDF.`,
              variant: "destructive",
            });
        }
      }
    }
  
    try {
      if (successfulPages > 0) {
        toast({
          title: "Preparando descarga...",
          description: "El archivo PDF se está compilando y se descargará en breve.",
        });
        pdf.save('reporte.pdf');
      } else {
        toast({
            title: "No se generó el PDF",
            description: "No se pudo procesar ningún reporte.",
            variant: "destructive",
        });
      }
    } catch (e) {
        console.error("Error saving PDF:", e);
        toast({
            title: "Error al guardar el PDF",
            description: "El documento es demasiado grande y no se pudo generar. Intenta con menos datos.",
            variant: "destructive",
        });
    } finally {
        setLoading(false);
        setProgress(0);
    }
  };
  
  const resetState = () => {
    setProcessedData(null);
    setFileName(null);
    if(fileInputRef.current) {
      fileInputRef.current.value = "";
    }
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
          Reportes de utilidad venta en verde
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube tu archivo de Excel para generar un reporte de utilidad de venta en verde en PDF listo para imprimir.
        </p>
      </div>

      {!processedData ? (
        <Card className="max-w-xl mx-auto shadow-lg border-2 border-dashed border-primary/50 hover:border-primary transition-colors">
          <CardContent className="p-8">
            {parsing ? (
              <div className="flex flex-col items-center justify-center space-y-4 text-center">
                  <Loader2 className="w-16 h-16 text-primary animate-spin" />
                  <p className="text-lg font-semibold text-foreground">Procesando archivo...</p>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center space-y-4 text-center">
                <UploadCloud className="w-16 h-16 text-primary" />
                <p className="text-lg font-semibold text-foreground">Arrastra y suelta tu archivo de Excel aquí</p>
                <p className="text-muted-foreground">o</p>
                <Button onClick={() => fileInputRef.current?.click()}>
                  Seleccionar Archivo
                </Button>
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  onChange={handleFileChange}
                  accept=".xlsx, .xls"
                />
              </div>
            )}
          </CardContent>
        </Card>
      ) : (
        <div className="space-y-8">
          <Card className="max-w-7xl mx-auto shadow-lg">
            <CardHeader className="flex flex-row items-center justify-between">
                <div>
                    <CardTitle className="font-headline text-2xl">Previsualización del Reporte</CardTitle>
                    <CardDescription>Archivo: {fileName}</CardDescription>
                </div>
                <div className="flex gap-2">
                    <Button variant="outline" onClick={resetState} disabled={loading}>
                      <FileX2 className="mr-2 h-4 w-4" />
                      Cargar Otro
                    </Button>
                    <Button onClick={handleDownloadPdf} disabled={loading}>
                      {loading ? (
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      ) : (
                        <FileDown className="mr-2 h-4 w-4" />
                      )}
                      {loading ? 'Generando...' : 'Descargar PDF'}
                    </Button>
                </div>
            </CardHeader>
            <CardContent>
                {loading && (
                    <div className="space-y-2 mb-4">
                        <Progress value={progress} />
                        <p className="text-sm text-center text-muted-foreground">
                            Generando PDF... {Math.round(progress)}%
                        </p>
                    </div>
                )}
                <div id="pdf-content" ref={pdfContentRef} className="space-y-8">
                {processedData.map((group) => (
                    <div key={group.n_doc} className="p-4 bg-white text-black">
                      <header className="mb-2">
                          <h2 className="font-bold text-primary text-base">Reporte utilidad venta en verde</h2>
                      </header>
                      <Table className="border-collapse text-[8px]">
                          <TableHeader>
                            <TableRow className="bg-primary/20 hover:bg-primary/20">
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Doc.mat.</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Factura</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Nº doc.</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Ce.</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Fecha Factura</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Proveedor</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Nombre del proveedor</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Material</TableHead>
                              <TableHead className="px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Texto breve de material</TableHead>
                              <TableHead className="text-right px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Cantidad</TableHead>
                              <TableHead className="text-right px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Costo Total</TableHead>
                              <TableHead className="text-right px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Precio Venta</TableHead>
                              <TableHead className="text-right px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Utilidad %</TableHead>
                              <TableHead className="text-right px-1 py-0.5 h-auto text-black font-bold border border-neutral-300">Valor a pagar</TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                          {group.items.map((item, itemIndex) => {
                              const date = new Date(item['Fecha Factura']);
                              const formattedDate = Number.isNaN(date.getTime()) ? '' : `${String(date.getDate()).padStart(2, '0')}.${String(date.getMonth() + 1).padStart(2, '0')}.${date.getFullYear()}`;
                              return (
                              <TableRow key={itemIndex}>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Documento material']}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Factura']}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{getDocId(item)}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Centro']}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{formattedDate}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Proveedor']}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Nombre del proveedor']}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Material']}</TableCell>
                                <TableCell className="px-1 py-0.5 border border-neutral-300">{item['Texto breve de material']}</TableCell>
                                <TableCell className="text-right px-1 py-0.5 border border-neutral-300">{(item['Cantidad'] ?? 0).toFixed(3)}</TableCell>
                                <TableCell className="text-right px-1 py-0.5 border border-neutral-300 font-medium text-[9px]">{(item['Costo Total'] ?? 0).toFixed(2)}</TableCell>
                                <TableCell className="text-right px-1 py-0.5 border border-neutral-300">{(item['Precio Venta'] ?? 0).toFixed(2)}</TableCell>
                                <TableCell className="text-right px-1 py-0.5 border border-neutral-300">{(item['Utilidad %'] ?? 0).toFixed(2)}</TableCell>
                                <TableCell className="text-right px-1 py-0.5 border border-neutral-300">{(item['Valor a pagar'] ?? 0).toFixed(2)}</TableCell>
                              </TableRow>
                              )
                          })}
                          </TableBody>
                          <TableFooter>
                            <TableRow className="bg-accent/30 hover:bg-accent/30 font-bold">
                              <TableCell className="text-left font-bold px-1 py-0.5 border border-neutral-300">*</TableCell>
                              <TableCell colSpan={8} className="px-1 py-0.5 border border-neutral-300"></TableCell>
                              <TableCell className="text-right font-bold px-1 py-0.5 border border-neutral-300">
                                {group.totalCantidad.toFixed(3)}
                              </TableCell>
                              <TableCell className="text-right font-bold px-1 py-0.5 border border-neutral-300 text-[9px]">
                                {group.totalCostoTotal.toFixed(2)}
                              </TableCell>
                              <TableCell className="text-right font-bold px-1 py-0.5 border border-neutral-300">
                                {group.totalPrecioVenta.toFixed(2)}
                              </TableCell>
                              <TableCell className="text-right font-bold px-1 py-0.5 border border-neutral-300">
                                {group.items.length > 0 ? (group.totalUtilidad / group.items.length).toFixed(2) : '0.00'}
                              </TableCell>
                              <TableCell className="text-right font-bold px-1 py-0.5 border border-neutral-300">
                                {group.totalValorAPagar.toFixed(2)}
                              </TableCell>
                            </TableRow>
                          </TableFooter>
                      </Table>
                    </div>
                ))}
                </div>
            </CardContent>
          </Card>
        </div>
      )}

      <section className="mt-16 max-w-4xl mx-auto">
        <h3 className="text-3xl font-headline font-bold text-center mb-8 text-foreground">
          ¿Cómo funciona?
        </h3>
        <div className="grid md:grid-cols-3 gap-8 text-center">
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">1</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Sube tu Archivo</h4>
            <p className="text-foreground/80">
              Arrastra o selecciona tu archivo de Excel (.xlsx o .xls) con los datos de la venta en verde.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">2</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Previsualiza el Reporte</h4>
            <p className="text-foreground/80">
              La aplicación agrupará los datos y te mostrará una vista previa de cómo se verá cada reporte en el PDF.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">3</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Descarga el PDF</h4>
            <p className="text-foreground/80">
              Haz clic en "Descargar PDF" para obtener un archivo listo para imprimir con todos los reportes, cada uno en su propia página.
            </p>
          </div>
        </div>
      </section>
    </main>
  );
}
