
"use client";

import { useState, useRef, ChangeEvent, useEffect } from 'react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, XCircle, ArrowLeft, CheckCircle, RefreshCw, FileText } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
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
    const exactKey = keys.find(key => key.trim().toLowerCase() === 'nº doc.');
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
        description: `No se encontraron datos para agrupar. Revisa que la columna 'Nº doc.' exista y tenga valores. Columnas encontradas: ${firstRowKeys}`,
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
                    (item['Cantidad'] ?? 0).toFixed(2),
                    `${(item['Utilidad %'] ?? 0).toFixed(0)} %`,
                    (item['Costo Total'] ?? 0).toFixed(2),
                    (item['Precio Venta'] ?? 0).toFixed(2),
                    (item['Valor a pagar'] ?? 0).toFixed(2)
                ];
            });

            const totalUtilidadAvg = group.items.length > 0 ? `${(group.totalUtilidad / group.items.length).toFixed(0)} %` : '0 %';

            const foot = [[
                '*',
                { content: '', colSpan: 8 },
                group.totalCantidad.toFixed(2),
                totalUtilidadAvg,
                group.totalCostoTotal.toFixed(2),
                group.totalPrecioVenta.toFixed(2),
                group.totalValorAPagar.toFixed(2)
            ]];

            autoTable(pdf, {
                startY: 20,
                margin: { left: 5 },
                head: head,
                body: body,
                foot: foot,
                theme: 'grid',
                headStyles: {
                    fillColor: [221, 237, 255],
                    textColor: 0,
                    fontSize: 5,
                    cellPadding: 1,
                    halign: 'center',
                    lineColor: [0, 0, 0],
                    lineWidth: 0.1,
                },
                bodyStyles: {
                    textColor: 0,
                    cellPadding: 1,
                    halign: 'center',
                    lineColor: [0, 0, 0],
                    lineWidth: 0.1,
                },
                footStyles: {
                    fillColor: [250, 235, 215],
                    textColor: 0,
                    fontSize: 6,
                    cellPadding: 1,
                    halign: 'center',
                    fontStyle: 'normal',
                    lineColor: [0, 0, 0],
                    lineWidth: 0.1,
                },
                columnStyles: {
                    0: { cellWidth: 17 },
                    1: { cellWidth: 25 },
                    2: { cellWidth: 17 },
                    3: { cellWidth: 10 },
                    4: { cellWidth: 15 },
                    5: { cellWidth: 16 },
                    6: { cellWidth: 47 },
                    7: { cellWidth: 26 },
                    8: { cellWidth: 47 },
                    9: { cellWidth: 12 },
                    10: { cellWidth: 12 },
                    11: { cellWidth: 13 },
                    12: { cellWidth: 15 },
                    13: { cellWidth: 15 },
                },
                didParseCell: function (data) {
                    if (data.row.section === 'body') {
                        if (data.column.index !== 6 && data.column.index !== 8) {
                            data.cell.styles.fontSize = 7;
                        } else {
                            data.cell.styles.fontSize = 5;
                        }
                    }
                    if (data.row.section === 'foot' && data.column.index === 11) {
                             data.cell.styles.fontStyle = 'bold';
                             data.cell.styles.fontSize = 7;
                        }
                }
            });
            
            await new Promise(resolve => setTimeout(resolve, 50));
            setProgress(((i + 1) / totalPages) * 100);
        }

        if (cancelPdfGeneration.current) {
          setStatus('idle');
          toast({ title: "Generación de PDF cancelada" });
          return;
        }
      
        if (processedData.length > 0) {
          pdf.save('reporte_venta_en_verde.pdf');
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
                <div className="flex flex-col items-center justify-center space-y-4 text-center animate-in fade-in zoom-in duration-300">
                    <div className="p-4 bg-primary/10 rounded-full mb-2">
                        <Loader2 className="w-10 h-10 text-primary animate-spin" />
                    </div>
                    <p className="text-xl font-semibold tracking-tight text-foreground">Procesando tu archivo Excel...</p>
                    <p className="text-sm text-muted-foreground">{fileName}</p>
                </div>
            );
        case 'generating':
            return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center animate-in fade-in zoom-in duration-300">
                  <div className="p-4 bg-primary/10 rounded-full mb-2 relative">
                      <div className="absolute inset-0 border-4 border-t-primary border-primary/20 rounded-full animate-spin"></div>
                      <FileDown className="w-8 h-8 text-primary opacity-50" />
                  </div>
                  <p className="text-xl font-semibold tracking-tight text-foreground">Generando PDF...</p>
                   <div className="w-full max-w-md pt-4 space-y-3">
                         <div className="h-2 bg-secondary rounded-full overflow-hidden">
                             <div className="h-full bg-gradient-to-r from-primary to-blue-500 transition-all duration-300 ease-out" style={{ width: `${progress}%` }}></div>
                         </div>
                         {processedData && <p className="text-xs font-medium text-muted-foreground uppercase tracking-wider">Página {Math.min(Math.ceil(progress * processedData.length / 100), processedData.length)} de {processedData.length}</p>}
                    </div>
                </div>
            );
        case 'success':
            return (
                <div className="flex flex-col items-center justify-center space-y-6 text-center animate-in fade-in zoom-in duration-500">
                    <div className="p-5 bg-emerald-500/10 rounded-full relative">
                        <div className="absolute inset-0 bg-emerald-500/20 rounded-full animate-ping"></div>
                        <CheckCircle className="w-12 h-12 text-emerald-500" />
                    </div>
                    <div>
                        <p className="text-2xl font-bold tracking-tight text-foreground">¡Reporte Generado con Éxito!</p>
                        <p className="text-sm text-muted-foreground mt-2">La descarga ha comenzado autómaticamente para: <br/><span className="font-medium text-foreground">{fileName}</span></p>
                    </div>
                    <Button onClick={resetState} className="mt-4 shadow-lg shadow-primary/20 transition-all hover:scale-105">
                        <RefreshCw className="mr-2 h-4 w-4" />
                        Generar Otro Reporte
                    </Button>
                </div>
            );
        case 'error':
             return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center animate-in fade-in zoom-in duration-300">
                    <div className="p-4 bg-destructive/10 rounded-full mb-2">
                        <XCircle className="w-12 h-12 text-destructive" />
                    </div>
                    <div>
                        <p className="text-xl font-bold tracking-tight text-foreground">Ocurrió un Error</p>
                        <p className="text-sm text-muted-foreground mt-1 max-w-sm">No se pudo generar el reporte. Por favor, revisa el archivo de Excel y vuelve a intentarlo.</p>
                    </div>
                    <Button variant="outline" onClick={resetState} className="mt-2 hover:bg-destructive hover:text-destructive-foreground transition-colors">
                        <RefreshCw className="mr-2 h-4 w-4" />
                        Intentar de Nuevo
                    </Button>
                </div>
            );
        case 'idle':
        default:
            return (
                <div className="flex flex-col items-center justify-center space-y-6 text-center">
                    <div className="w-24 h-24 rounded-full bg-primary/5 flex items-center justify-center mb-2 group-hover:scale-110 transition-transform duration-500">
                        <UploadCloud className="w-12 h-12 text-primary/80" strokeWidth={1.5} />
                    </div>
                    <div>
                        <p className="text-xl font-semibold tracking-tight text-foreground">Haz clic para subir tu archivo Excel</p>
                        <p className="text-sm text-muted-foreground mt-2">El archivo PDF se generará y descargará al instante</p>
                    </div>
                    <Button onClick={() => fileInputRef.current?.click()} size="lg" className="shadow-lg shadow-primary/20 hover:-translate-y-0.5 transition-all w-full sm:w-auto">
                        <FileDown className="mr-2 h-5 w-5" />
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
    <div className="min-h-screen bg-background relative overflow-hidden selection:bg-primary/30 selection:text-primary-foreground">
      {/* Sleek background effect: Grid + Glow */}
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:24px_24px] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_0%,#000_70%,transparent_100%)] pointer-events-none" />
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[1000px] h-[400px] opacity-30 dark:opacity-20 pointer-events-none [mask-image:radial-gradient(ellipse_at_center,black,transparent_80%)]">
        <div className="absolute inset-0 bg-gradient-to-r from-emerald-500/30 via-primary/40 to-blue-500/40 blur-[100px]" />
      </div>

      <div className="relative mx-auto max-w-5xl px-6 py-12 lg:px-8 z-10 flex flex-col min-h-screen">
        <div className="mb-10 flex justify-center sm:justify-start">
            <Link href="/" className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-muted/40 border border-border/50 text-sm font-medium text-muted-foreground backdrop-blur-sm shadow-sm hover:bg-muted/60 hover:text-foreground transition-all">
                <ArrowLeft className="w-4 h-4" />
                Volver al Workspace
            </Link>
        </div>

        <header className="flex flex-col items-center text-center max-w-3xl mx-auto mb-16 space-y-4">
            <div className="inline-flex items-center justify-center p-3 rounded-2xl bg-emerald-500/10 text-emerald-500 ring-1 ring-emerald-500/20 mb-2">
                <FileText className="w-6 h-6" strokeWidth={1.5} />
            </div>
            <h1 className="text-4xl sm:text-5xl font-bold tracking-tight text-foreground">
                Venta en <span className="bg-gradient-to-br from-emerald-500 to-emerald-700 bg-clip-text text-transparent">Verde</span>
            </h1>
            <p className="text-lg text-muted-foreground max-w-xl">
                Sube tu reporte resumido en Excel y obtén los reportes individuales listos para imprimir en PDF de forma automática.
            </p>
        </header>

        <main className="flex-grow w-full max-w-3xl mx-auto">
            <Card className="relative overflow-hidden border border-border/40 bg-background/40 backdrop-blur-xl shadow-2xl shadow-primary/5 transition-all">
                {/* Glowing border effect at the top */}
                <div className="absolute top-0 inset-x-0 h-[1px] bg-gradient-to-r from-transparent via-primary/50 to-transparent"></div>
                <CardContent className="p-8 sm:p-12 min-h-[350px] flex items-center justify-center group">
                    {renderStatus()}
                </CardContent>
            </Card>
            
            <div className="grid md:grid-cols-2 gap-6 mt-16">
                <section className="col-span-1 md:col-span-2">
                    <h3 className="text-2xl font-bold tracking-tight mb-6 flex items-center gap-2">
                        <div className="w-8 h-8 rounded-lg bg-primary/10 flex items-center justify-center text-primary">
                            <span className="font-semibold text-sm">1</span>
                        </div>
                        ¿Cómo funciona?
                    </h3>
                    <div className="grid sm:grid-cols-3 gap-4">
                        <div className="p-5 rounded-2xl bg-card/30 backdrop-blur-sm border border-border/40">
                            <UploadCloud className="w-6 h-6 text-primary/70 mb-3" />
                            <h4 className="font-semibold mb-1">Sube el Archivo</h4>
                            <p className="text-sm text-muted-foreground leading-relaxed">Carga tu exportación de SAP en formato Excel.</p>
                        </div>
                        <div className="p-5 rounded-2xl bg-card/30 backdrop-blur-sm border border-border/40">
                            <RefreshCw className="w-6 h-6 text-primary/70 mb-3" />
                            <h4 className="font-semibold mb-1">Cálculo Preciso</h4>
                            <p className="text-sm text-muted-foreground leading-relaxed">Agrupamos y calculamos automáticamente los totales.</p>
                        </div>
                        <div className="p-5 rounded-2xl bg-card/30 backdrop-blur-sm border border-border/40">
                            <CheckCircle className="w-6 h-6 text-emerald-500/70 mb-3" />
                            <h4 className="font-semibold mb-1">PDF Listo</h4>
                            <p className="text-sm text-muted-foreground leading-relaxed">El documento formateado se descarga sin demoras.</p>
                        </div>
                    </div>
                </section>

                <section className="col-span-1 md:col-span-2 mt-8">
                    <div className="p-6 sm:p-8 rounded-3xl bg-muted/20 backdrop-blur-xl border border-border/40">
                        <h3 className="text-xl font-bold text-foreground mb-6 flex items-center gap-3">
                            Extracción desde SAP
                            <span className="text-xs px-2 py-1 bg-primary/10 text-primary rounded-full font-medium tracking-wide">GUÍA</span>
                        </h3>
                        <ol className="relative border-s border-border/60 ml-3 space-y-6">
                            <li className="ms-6">
                                <span className="absolute flex items-center justify-center w-6 h-6 bg-background rounded-full -start-3 ring-4 ring-background border border-border/80 text-xs font-bold text-muted-foreground">1</span>
                                <p className="text-sm text-foreground/90 leading-relaxed">Ejecuta la transacción <strong className="text-primary font-mono bg-primary/5 px-1 py-0.5 rounded">ZMM_UTILIDAD_VV</strong> en el sistema SAP.</p>
                            </li>
                            <li className="ms-6">
                                <span className="absolute flex items-center justify-center w-6 h-6 bg-background rounded-full -start-3 ring-4 ring-background border border-border/80 text-xs font-bold text-muted-foreground">2</span>
                                <p className="text-sm text-foreground/90 leading-relaxed">Configura el <strong>rango de fechas</strong> a evaluar y define el rango de documentos en "Número de documento" (ej. de 52000xxxx0 a 52000xxxx9).</p>
                            </li>
                            <li className="ms-6">
                                <span className="absolute flex items-center justify-center w-6 h-6 bg-background rounded-full -start-3 ring-4 ring-background border border-border/80 text-xs font-bold text-muted-foreground">3</span>
                                <p className="text-sm text-foreground/90 leading-relaxed">Presiona <strong className="bg-muted px-1.5 py-0.5 rounded shadow-sm text-xs border border-border">F8</strong> o clic en Ejecutar para generar la tabla de resultados.</p>
                            </li>
                            <li className="ms-6">
                                <span className="absolute flex items-center justify-center w-6 h-6 bg-background rounded-full -start-3 ring-4 ring-background border border-border/80 text-xs font-bold text-muted-foreground">4</span>
                                <p className="text-sm text-foreground/90 leading-relaxed">Ve al menú <strong className="font-medium">Exportar &gt; Hoja de Cálculo</strong>, guarda el archivo como Excel y súbelo aquí arriba.</p>
                            </li>
                        </ol>
                    </div>
                </section>
            </div>
        </main>
      </div>
    </div>
  );
}
