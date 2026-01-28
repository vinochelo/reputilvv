
"use client";

import { useState, useRef, ChangeEvent, useEffect } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, ArrowLeft, CheckCircle, XCircle, RefreshCw, Shuffle } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow, TableFooter } from '@/components/ui/table';
import { Progress } from '@/components/ui/progress';
import { useToast } from '@/hooks/use-toast';
import Link from 'next/link';

type Status = 'idle' | 'parsing' | 'generating' | 'success' | 'error';
type RetailData = { [key: string]: any };
type GroupedData = {
  n_doc: number | string;
  items: RetailData[];
  totalCantidad: number;
  totalImporte: number;
};

// Helper to get a value from an object with multiple possible keys, ignoring case, spaces, and dots.
const getValue = (item: any, keys: string[]): any => {
    const itemKeys = Object.keys(item);
    for (const key of keys) {
        const normalizedKey = key.toLowerCase().replace(/[\.\s]/g, '');
        const foundKey = itemKeys.find(ik => ik.trim().toLowerCase().replace(/[\.\s]/g, '') === normalizedKey);
        if (foundKey && item[foundKey] !== null && item[foundKey] !== undefined) {
            return item[foundKey];
        }
    }
    return undefined;
};

// Accessor functions for specific fields
const getDocId = (item: RetailData) => getValue(item, ['nºdocumento', 'belnr', 'nro.documento']);
const getOrder = (item: RetailData) => getValue(item, ['ord.decompra', 'ebeln']);
const getProvider = (item: RetailData) => getValue(item, ['proveedor', 'lifnr']);
const getProviderName = (item: RetailData) => getValue(item, ['nombredelproveedor']);
const getMaterial = (item: RetailData) => getValue(item, ['material']);
const getMaterialText = (item: RetailData) => getValue(item, ['textobrevedematerial']);
const getQuantity = (item: RetailData) => parseFloat(getValue(item, ['cantidad', 'menge'])) || 0;
const getAmount = (item: RetailData) => parseFloat(getValue(item, ['importeenmon.local', 'wrbtr'])) || 0;
const getPostingDate = (item: RetailData) => getValue(item, ['fechacontab.', 'bldat']);


export default function ReporteRetailPage() {
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

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(file.type)) {
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
          const jsonData = XLSX.utils.sheet_to_json<RetailData>(worksheet);
          
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
        toast({ title: "Error al leer el archivo", variant: "destructive" });
        setStatus('error');
      }
      reader.readAsArrayBuffer(file);
    }
  };

  const processData = (data: RetailData[]): GroupedData[] | null => {
    const grouped = data.reduce<{[key: string]: GroupedData}>((acc, item) => {
      const docId = getDocId(item);
      if (docId === undefined || docId === null) return acc;

      if (!acc[docId]) {
        acc[docId] = {
          n_doc: docId,
          items: [],
          totalCantidad: 0,
          totalImporte: 0,
        };
      }
      
      const cantidad = getQuantity(item);
      const importe = getAmount(item);
      
      acc[docId].items.push(item);
      acc[docId].totalCantidad += cantidad;
      acc[docId].totalImporte += importe;

      return acc;
    }, {});

    const groupedArray = Object.values(grouped);

    if (groupedArray.length === 0 && data.length > 0) {
      const firstRowKeys = Object.keys(data[0]).join(', ');
      toast({
        title: "No se pudo agrupar",
        description: `No se encontraron datos para agrupar. Revisa que la columna 'Nº documento' o 'BELNR' exista. Columnas encontradas: ${firstRowKeys}`,
        variant: "destructive",
        duration: 9000,
      });
      return null;
    }

    // Sort by document number
    groupedArray.sort((a, b) => String(a.n_doc).localeCompare(String(b.n_doc), undefined, { numeric: true }));
    return groupedArray;
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
            
            await new Promise(resolve => setTimeout(resolve, 10));
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
        pdf.save('reporte_retail.pdf');
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
                        <Shuffle className="mr-2 h-4 w-4" />
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
          Reportes de Retail
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube tu reporte de SAP en Excel para generar un PDF para imprimir, ordenado y agrupado por número de documento.
        </p>
      </div>

      <Card className="max-w-2xl mx-auto shadow-lg">
          <CardContent className="p-8 min-h-[250px] flex items-center justify-center">
            {renderStatus()}
          </CardContent>
      </Card>
      
      <div className="absolute -left-[9999px] top-0 opacity-0" ref={pdfContentRef}>
        {processedData?.map((group) => (
            <div key={group.n_doc} className="p-4 bg-white text-black w-[1123px]">
              <header className="mb-2">
                  <h2 className="text-primary text-base font-normal">Reporte Retail - Documento: {group.n_doc}</h2>
              </header>
              <Table className="border-collapse text-[8px]">
                  <TableHeader>
                    <TableRow className="bg-primary/20 hover:bg-primary/20">
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Ord. de Compra</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Proveedor</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Nombre Proveedor</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Material</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Texto de Material</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Fecha Contab.</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Cantidad</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Importe</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                  {group.items.map((item, itemIndex) => {
                      let formattedDate = '';
                      const dateValue = getPostingDate(item);
                      if (dateValue) {
                        try {
                           const date = new Date(dateValue);
                           if (!Number.isNaN(date.getTime())) {
                             // Adjust for timezone offset
                             const adjustedDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60 * 1000);
                             formattedDate = `${String(adjustedDate.getDate()).padStart(2, '0')}.${String(adjustedDate.getMonth() + 1).padStart(2, '0')}.${adjustedDate.getFullYear()}`;
                           }
                        } catch(e) { /* Ignore date parsing errors */ }
                      }
                      
                      return (
                      <TableRow key={itemIndex}>
                        <TableCell className="px-1 py-1 border border-neutral-300">{getOrder(item)}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{getProvider(item)}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{getProviderName(item)}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{getMaterial(item)}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{getMaterialText(item)}</TableCell>
                        <TableCell className="px-1 py-1 border border-neutral-300">{formattedDate}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300">{getQuantity(item)?.toFixed(3)}</TableCell>
                        <TableCell className="text-right px-1 py-1 border border-neutral-300 font-medium text-[9px]">{getAmount(item)?.toFixed(2)}</TableCell>
                      </TableRow>
                      )
                  })}
                  </TableBody>
                  <TableFooter>
                    <TableRow className="bg-accent/30 hover:bg-accent/30 font-bold">
                      <TableCell colSpan={6} className="px-1 py-1 border border-neutral-300 text-left">TOTAL</TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300">
                        {group.totalCantidad.toFixed(3)}
                      </TableCell>
                      <TableCell className="text-right font-bold px-1 py-1 border border-neutral-300 text-[9px]">
                        {group.totalImporte.toFixed(2)}
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
