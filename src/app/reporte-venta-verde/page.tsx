
"use client";

import { useState, useRef, ChangeEvent, useEffect } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, FileX2, ArrowLeft, CheckCircle, XCircle, RefreshCw } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow, TableFooter } from '@/components/ui/table';
import { Progress } from '@/components/ui/progress';
import { useToast } from '@/hooks/use-toast';
import type { ExcelData, GroupedData } from '@/lib/types';
import Link from 'next/link';

type Status = 'idle' | 'parsing' | 'generating' | 'success' | 'error';

export default function ReporteVentaVerdePage() {
  const [processedData, setProcessedData] = useState<GroupedData[] | null>(null);
  const [status, setStatus] = useState<Status>('idle');
  const [fileName, setFileName] = useState<string | null>(null);
  const [progress, setProgress] = useState(0);
  const pdfContentRef = useRef<HTMLDivElement>(null);
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const cancelPdfGeneration = useRef(false);

  useEffect(() => {
    const generate = async () => {
      if (processedData && processedData.length > 0 && status === 'parsing') {
        await handleDownloadPdf();
      }
    };
    generate();
  }, [processedData, status]);


  const getDocId = (item: any): number | string | undefined => {
    const keys = Object.keys(item);
    const exactKey = keys.find(key => key.trim().toLowerCase() === 'nº documento');
    if (exactKey && item[exactKey] !== null && item[exactKey] !== undefined) {
        return item[exactKey];
    }
    
    const docIdRegex = /n(º|°|o|ro\.?|umero)?\.?\s*doc(umento)?/i;
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
      resetState();
      setFileName(file.name);
      setStatus('parsing');

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json<ExcelData>(worksheet);
          
          const groupedData = processData(jsonData);
          if (groupedData && groupedData.length > 0) {
            setProcessedData(groupedData);
          } else {
            setStatus('error');
          }
        } catch (error) {
          console.error("Error parsing Excel file:", error);
          toast({
            title: "Error al procesar el archivo",
            description: "No se pudo leer el archivo de Excel. Asegúrate de que el formato sea correcto.",
            variant: "destructive",
          });
          setStatus('error');
        }
      };
      reader.onerror = () => {
        console.error("Error reading file");
        toast({
            title: "Error al leer el archivo",
            description: "Ocurrió un error al intentar leer el archivo.",
            variant: "destructive",
        });
        setStatus('error');
      }
      reader.readAsArrayBuffer(file);
    }
  };

  const processData = (data: ExcelData[]): GroupedData[] | null => {
    const grouped = data.reduce<GroupedData[]>((acc, item) => {
      const docId = getDocId(item);

      if (docId === undefined || docId === null) {
        return acc;
      }

      let group = acc.find(g => g.n_doc === docId);

      if (!group) {
        group = {
          n_doc: docId as number,
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
        description: `No se encontraron datos para agrupar. Revisa que la columna 'Nº documento' exista y tenga valores. Columnas encontradas: ${firstRowKeys}`,
        variant: "destructive",
      });
      return null;
    }

    grouped.sort((a, b) => a.n_doc - b.n_doc);
    return grouped;
  };

  const handleDownloadPdf = async () => {
    if (!pdfContentRef.current) {
        setStatus('error');
        toast({ title: "Error", description: "No se encontró el contenido para generar el PDF.", variant: "destructive" });
        return;
    };
  
    setStatus('generating');
    setProgress(0);
    cancelPdfGeneration.current = false;
    await new Promise(resolve => setTimeout(resolve, 50));
    
    const pdf = new jsPDF('l', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const reportElements = Array.from(pdfContentRef.current.children);
    let successfulPages = 0;
    
    const margin = 5;
    const imageWidth = pdfWidth - (margin * 2);

    for (let i = 0; i < reportElements.length; i++) {
      if (cancelPdfGeneration.current) break;
      const reportElement = reportElements[i] as HTMLElement;
      if (reportElement) {
        try {
            const canvas = await html2canvas(reportElement, { scale: 2, useCORS: true });
            const imgData = canvas.toDataURL('image/jpeg', 0.95); 
            
            if (i > 0) pdf.addPage();
  
            const imgProps = pdf.getImageProperties(imgData);
            const imageHeight = (imgProps.height * imageWidth) / imgProps.width;
            
            pdf.addImage(imgData, 'JPEG', margin, margin, imageWidth, imageHeight);
            successfulPages++;
            setProgress(((i + 1) / reportElements.length) * 100);
            
            await new Promise(resolve => setTimeout(resolve, 10)); // Shorter delay
        } catch(e) {
            console.error(`Error processing page ${i + 1} for PDF:`, e);
        }
      }
    }

    if (cancelPdfGeneration.current) {
      setStatus('idle');
      toast({ title: "Generación de PDF cancelada" });
      return;
    }
  
    if (successfulPages > 0) {
      try {
        pdf.save('reporte_venta_en_verde.pdf');
        setStatus('success');
        toast({ title: "¡Éxito! PDF generado.", description: "La descarga de tu archivo ha comenzado." });
      } catch (e) {
        console.error("Error saving PDF:", e);
        toast({ title: "Error al guardar el PDF", description: "El documento es demasiado grande. Intenta con menos datos.", variant: "destructive" });
        setStatus('error');
      }
    } else {
        toast({ title: "No se generó el PDF", description: "No se pudo procesar ningún reporte.", variant: "destructive" });
        setStatus('error');
    }
  };

  const handleCancelPdfGeneration = () => {
    cancelPdfGeneration.current = true;
  };
  
  const resetState = () => {
    setProcessedData(null);
    setFileName(null);
    setStatus('idle');
    setProgress(0);
    cancelPdfGeneration.current = false;
    if(fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const renderStatus = () => {
    switch (status) {
        case 'parsing':
            return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <Loader2 className="w-12 h-12 text-primary animate-spin" />
                    <p className="text-lg font-semibold text-foreground">Procesando tu archivo Excel...</p>
                    <p className="text-sm text-muted-foreground">{fileName}</p>
                </div>
            );
        case 'generating':
            return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <Loader2 className="w-16 h-16 text-primary animate-spin" />
                    <p className="text-lg font-semibold text-foreground">Generando PDF, por favor espera...</p>
                    <div className="w-full space-y-4">
                        <Progress value={progress} className="h-6" />
                        <p className="text-3xl text-center font-bold text-primary animate-pulse tabular-nums">
                            {Math.round(progress)}%
                        </p>
                    </div>
                    <Button variant="destructive" size="sm" onClick={handleCancelPdfGeneration}>
                        <XCircle className="mr-2 h-4 w-4" />
                        Cancelar
                    </Button>
                </div>
            );
        case 'success':
            return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <CheckCircle className="w-12 h-12 text-green-600" />
                    <p className="text-lg font-semibold text-foreground">¡PDF Generado con Éxito!</p>
                    <p className="text-sm text-muted-foreground">Tu descarga ha comenzado para: {fileName}</p>
                    <Button onClick={resetState}>
                        <RefreshCw className="mr-2 h-4 w-4" />
                        Generar Otro Reporte
                    </Button>
                </div>
            );
        case 'error':
             return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <XCircle className="w-12 h-12 text-destructive" />
                    <p className="text-lg font-semibold text-foreground">Ocurrió un Error</p>
                    <p className="text-sm text-muted-foreground">No se pudo generar el reporte. Por favor, revisa el archivo e inténtalo de nuevo.</p>
                    <Button variant="outline" onClick={resetState}>
                        <RefreshCw className="mr-2 h-4 w-4" />
                        Intentar de Nuevo
                    </Button>
                </div>
            );
        case 'idle':
        default:
            return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <UploadCloud className="w-16 h-16 text-primary" />
                    <p className="text-lg font-semibold text-foreground">Haz clic para subir tu archivo de Excel</p>
                    <p className="text-muted-foreground">El PDF se generará y descargará al instante</p>
                    <Button onClick={() => fileInputRef.current?.click()}>
                        <FileDown className="mr-2 h-4 w-4" />
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
            );
    }
  }

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
          Sube tu archivo de Excel para generar un reporte PDF listo para imprimir. La descarga comenzará automáticamente.
        </p>
      </div>

      <Card className="max-w-2xl mx-auto shadow-lg">
          <CardContent className="p-8 min-h-[250px] flex items-center justify-center">
            {renderStatus()}
          </CardContent>
      </Card>

      <section className="mt-16 max-w-4xl mx-auto">
        <h3 className="text-3xl font-headline font-bold text-center mb-8 text-foreground">
          Instrucciones para obtener el Excel de SAP
        </h3>
        <div className="bg-muted/50 p-6 rounded-lg">
            <ol className="list-decimal list-inside space-y-4 text-foreground/90">
                <li>
                    Ingresa a SAP y ejecuta la transacción <strong>ZMM_UTILIDAD_VV</strong>.
                </li>
                <li>
                    En el campo de selección de fechas, coloca el <strong>rango de fechas</strong> correspondiente a las facturas que deseas procesar.
                </li>
                <li>
                    En la sección <strong>"Número de documento"</strong>, especifica el rango de los documentos FI. Debes usar el primer y el último número de documento del período (ej: desde 52000xxxx0 hasta 52000xxxx9).
                </li>
                <li>
                    Haz clic en <strong>Ejecutar</strong> (o presiona F8).
                </li>
                <li>
                    Una vez que se muestren los resultados, ve al menú <strong>Exportar &gt; Hoja de Cálculo</strong>.
                </li>
                <li>
                    Guarda el archivo en formato Excel, asígnale un nombre descriptivo y ¡listo! Ya puedes subirlo a esta herramienta.
                </li>
            </ol>
        </div>
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
            <h4 className="text-xl font-semibold text-foreground">Sube tu Archivo</h4>
            <p className="text-foreground/80">
              Arrastra o selecciona tu archivo de Excel (.xlsx o .xls) con los datos de la venta en verde.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">2</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Procesamiento Automático</h4>
            <p className="text-foreground/80">
              La aplicación procesa los datos y genera un PDF al instante, sin pasos intermedios.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">3</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Descarga Instantánea</h4>
            <p className="text-foreground/80">
              Se inicia automáticamente la descarga de un archivo PDF listo para imprimir con todos tus reportes.
            </p>
          </div>
        </div>
      </section>
      
      {/* Hidden container for html2canvas to render the PDF content */}
      <div className="absolute -left-[9999px] top-0 opacity-0" ref={pdfContentRef}>
        {processedData?.map((group) => (
            <div key={group.n_doc} className="p-4 bg-white text-black">
              <header className="mb-2">
                  <h2 className="text-primary text-base font-normal">Reporte utilidad venta en verde</h2>
              </header>
              <Table className="border-collapse text-[8px]">
                  <TableHeader>
                    <TableRow className="bg-primary/20 hover:bg-primary/20">
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Doc.mat.</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Factura</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Nº doc.</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Ce.</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Fecha Factura</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Proveedor</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Nombre del proveedor</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Material</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Texto breve de material</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Cantidad</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Costo Total</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Precio Venta</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Utilidad %</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Valor a pagar</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                  {group.items.map((item, itemIndex) => {
                      const date = new Date(item['Fecha Factura']);
                      const formattedDate = Number.isNaN(date.getTime()) ? '' : `${String(date.getDate()).padStart(2, '0')}.${String(date.getMonth() + 1).padStart(2, '0')}.${date.getFullYear()}`;
                      return (
                      <TableRow key={itemIndex}>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Documento material']}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Factura']}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{getDocId(item)}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Centro']}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{formattedDate}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Proveedor']}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Nombre del proveedor']}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Material']}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{item['Texto breve de material']}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300">{(item['Cantidad'] ?? 0).toFixed(3)}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300 font-medium text-[9px]">{(item['Costo Total'] ?? 0).toFixed(2)}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300">{(item['Precio Venta'] ?? 0).toFixed(2)}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300">{(item['Utilidad %'] ?? 0).toFixed(2)}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300">{(item['Valor a pagar'] ?? 0).toFixed(2)}</TableCell>
                      </TableRow>
                      )
                  })}
                  </TableBody>
                  <TableFooter>
                    <TableRow className="bg-accent/30 hover:bg-accent/30 font-bold">
                      <TableCell className="text-left font-bold px-1 py-1 border border-neutral-300">*</TableCell>
                      <TableCell colSpan={8} className="px-1 py-1 border border-neutral-300"></TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300">
                        {group.totalCantidad.toFixed(3)}
                      </TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300 text-[9px]">
                        {group.totalCostoTotal.toFixed(2)}
                      </TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300">
                        {group.totalPrecioVenta.toFixed(2)}
                      </TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300">
                        {group.items.length > 0 ? (group.totalUtilidad / group.items.length).toFixed(2) : '0.00'}
                      </TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300">
                        {group.totalValorAPagar.toFixed(2)}
                      </TableCell>
                    </TableRow>
                  </TableFooter>
              </Table>
            </div>
        ))}
      </div>
    </main>
  );
}
