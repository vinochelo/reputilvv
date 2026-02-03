"use client";

import { useState, useRef, ChangeEvent, useEffect } from 'react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, XCircle, ArrowLeft, CheckCircle, RefreshCw } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Progress } from '@/components/ui/progress';
import { useToast } from '@/hooks/use-toast';
import type { ExcelData, GroupedData } from '@/lib/types';
import Link from 'next/link';

type Status = 'idle' | 'parsing' | 'generating' | 'success' | 'error';

export default function ReporteVentaVerdeBetaPage() {
  const [processedData, setProcessedData] = useState<GroupedData[] | null>(null);
  const [status, setStatus] = useState<Status>('idle');
  const [fileName, setFileName] = useState<string | null>(null);
  const [progress, setProgress] = useState(0);
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
    if (!processedData) {
        setStatus('error');
        toast({ title: "Error", description: "No hay datos procesados para generar el PDF.", variant: "destructive" });
        return;
    };
  
    setStatus('generating');
    setProgress(0);
    cancelPdfGeneration.current = false;
    await new Promise(resolve => setTimeout(resolve, 50));
    
    try {
        const pdf = new jsPDF('l', 'mm', 'a4');
        const totalPages = processedData.length;

        for (let i = 0; i < totalPages; i++) {
            if (cancelPdfGeneration.current) break;

            const group = processedData[i];
            if (i > 0) pdf.addPage();
            
            pdf.setFont('helvetica', 'normal');
            pdf.setFontSize(12);
            pdf.text('Reporte utilidad venta en verde', 14, 15);
            
            const head = [['Doc.mat.', 'Factura', 'Nº doc.', 'Ce.', 'Fecha Factura', 'Proveedor', 'Nombre del proveedor', 'Material', 'Texto breve de material', 'Cantidad', 'Utilidad %', 'Costo Total', 'Precio Venta', 'Valor a pagar']];
            
            const body = group.items.map(item => {
                const date = new Date(item['Fecha Factura']);
                const formattedDate = Number.isNaN(date.getTime()) ? '' : `${String(date.getDate()).padStart(2, '0')}.${String(date.getMonth() + 1).padStart(2, '0')}.${date.getFullYear()}`;
                return [
                    item['Documento material'],
                    item['Factura'],
                    getDocId(item),
                    item['Centro'],
                    formattedDate,
                    item['Proveedor'],
                    item['Nombre del proveedor'],
                    item['Material'],
                    item['Texto breve de material'],
                    (item['Cantidad'] ?? 0).toFixed(3),
                    (item['Utilidad %'] ?? 0).toFixed(2),
                    (item['Costo Total'] ?? 0).toFixed(2),
                    (item['Precio Venta'] ?? 0).toFixed(2),
                    (item['Valor a pagar'] ?? 0).toFixed(2)
                ];
            });

            const totalUtilidadAvg = group.items.length > 0 ? (group.totalUtilidad / group.items.length).toFixed(2) : '0.00';

            const foot = [['*', '', '', '', '', '', '', '', '', group.totalCantidad.toFixed(3), totalUtilidadAvg, { content: group.totalCostoTotal.toFixed(2), styles: { fontStyle: 'bold' } }, group.totalPrecioVenta.toFixed(2), group.totalValorAPagar.toFixed(2)]];

            autoTable(pdf, {
                startY: 20,
                head: head,
                body: body,
                foot: foot,
                theme: 'grid',
                headStyles: {
                    fillColor: [221, 237, 255],
                    textColor: 20,
                    fontSize: 5,
                    cellPadding: 1,
                },
                bodyStyles: {
                    fontSize: 5,
                    cellPadding: 1,
                },
                footStyles: {
                    fillColor: [250, 235, 215],
                    textColor: 20,
                    fontSize: 6,
                    cellPadding: 1,
                    fontStyle: 'normal',
                },
                columnStyles: {
                    9: { halign: 'right' },
                    10: { halign: 'right' },
                    11: { halign: 'right' },
                    12: { halign: 'right' },
                    13: { halign: 'right' },
                },
                didDrawPage: (data) => {
                    const str = "Página " + pdf.internal.getNumberOfPages();
                    pdf.setFontSize(10);
                    const pageSize = pdf.internal.pageSize;
                    const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();
                    pdf.text(str, data.settings.margin.left, pageHeight - 10);
                }
            });

            setProgress(((i + 1) / totalPages) * 100);
            await new Promise(resolve => setTimeout(resolve, 10));
        }

        if (cancelPdfGeneration.current) {
          setStatus('idle');
          toast({ title: "Generación de PDF cancelada" });
          return;
        }
      
        if (processedData.length > 0) {
          pdf.save('reporte_venta_en_verde_beta.pdf');
          setStatus('success');
          toast({ title: "¡Éxito! PDF generado.", description: "La descarga de tu archivo ha comenzado." });
        } else {
            setStatus('error');
        }

    } catch (e: any) {
        console.error("Error generating PDF with autoTable:", e);
        toast({ title: "Error al generar el PDF", description: e.message || "Ocurrió un error inesperado.", variant: "destructive" });
        setStatus('error');
    }
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
                    <Button variant="destructive" size="sm" onClick={() => (cancelPdfGeneration.current = true)}>
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
          Reportes Venta en Verde (Beta Rápido)
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Versión ultra-rápida. Genera el PDF directamente desde los datos, sin renderizar HTML.
        </p>
      </div>

      <Card className="max-w-2xl mx-auto shadow-lg">
          <CardContent className="p-8 min-h-[250px] flex items-center justify-center">
            {renderStatus()}
          </CardContent>
      </Card>
    </main>
  );
}
