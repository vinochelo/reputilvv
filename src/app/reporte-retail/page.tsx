
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
  n_doc: number | string; // This will hold the Document Number (BELNR)
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

// Accessor functions for specific fields, now more robust
const getOrder = (item: RetailData) => getValue(item, ['ord. de compra', 'ebeln', 'orden de compra']);
const getProvider = (item: RetailData) => getValue(item, ['proveedor', 'lifnr']);
const getProviderName = (item: RetailData) => getValue(item, ['nombre proveedor', 'nombredelproveedor']);
const getMaterial = (item: RetailData) => getValue(item, ['material']);
const getMaterialText = (item: RetailData) => getValue(item, ['descripcion material', 'textobrevedematerial']);
const getQuantity = (item: RetailData) => parseFloat(getValue(item, ['cant. in.', 'cantidad', 'menge'])) || 0;
const getAmount = (item: RetailData) => parseFloat(getValue(item, ['costo total.', 'costo total', 'importeenmon.local', 'wrbtr'])) || 0;
const getPostingDate = (item: RetailData) => getValue(item, ['fecha ingreso', 'fechacontab.', 'bldat']);

const readMainDataFile = (file: File): Promise<any[]> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = e.target?.result;
                if (!data) return reject(new Error(`El archivo de datos ${file.name} está vacío.`));
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                if (!sheetName) return reject(new Error(`No se encontraron hojas en ${file.name}.`));
                const worksheet = workbook.Sheets[sheetName];
                
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 3 });

                if (jsonData.length === 0) {
                    return reject(new Error(`No se encontraron filas de datos en ${file.name}. Por favor, asegúrate de que los datos comiencen en la fila 4.`));
                }

                resolve(jsonData);
            } catch (err: any) {
                reject(new Error(`Error al procesar el archivo de datos ${file.name}: ${err.message}`));
            }
        };
        reader.onerror = () => reject(new Error(`Error al leer el archivo ${file.name}.`));
        reader.readAsArrayBuffer(file);
    });
};

const readOrderFileAsHTML = (file: File): Promise<any[]> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const htmlContent = e.target?.result as string;
                if (!htmlContent) {
                    return reject(new Error(`El archivo de orden ${file.name} está vacío.`));
                }

                const tableContentMatch = htmlContent.match(/<table[\s\S]*?>([\s\S]*)<\/table>/i);
                if (!tableContentMatch) {
                    return reject(new Error("No se pudo encontrar una etiqueta <table> en el archivo de orden."));
                }

                const rowsHtml = tableContentMatch[1].match(/<tr[\s\S]*?>([\s\S]*?)<\/tr>/gi);
                if (!rowsHtml) {
                    return reject(new Error("No se encontraron filas (<tr>) en la tabla del archivo de orden."));
                }

                const allRows = rowsHtml.map(rowHtml => {
                    const cellsHtml = rowHtml.match(/<td[\s\S]*?>([\s\S]*?)<\/td>/gi);
                    return cellsHtml ? cellsHtml.map(cellHtml => cellHtml.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim()) : [];
                });

                let headerRowIndex = -1;
                let headerMap: { [key: string]: number } = {};
                for (let i = 0; i < allRows.length; i++) {
                    const cells = allRows[i].map(c => c.toUpperCase());
                    const ebelnIndex = cells.findIndex(c => c.includes("EBELN"));
                    const belnrIndex = cells.findIndex(c => c.includes("BELNR"));
                    if (ebelnIndex !== -1 && belnrIndex !== -1) {
                        headerRowIndex = i;
                        headerMap['ebeln'] = ebelnIndex;
                        headerMap['belnr'] = belnrIndex;
                        break;
                    }
                }

                if (headerRowIndex === -1) {
                    return reject(new Error("No se pudo encontrar la fila de encabezado con 'EBELN' y 'BELNR' en el archivo de orden."));
                }
                
                const dataRows = allRows.slice(headerRowIndex + 1);
                const jsonData = dataRows.map(rowCells => {
                    const ebeln = rowCells[headerMap['ebeln']] || "";
                    const belnr = rowCells[headerMap['belnr']] || "";
                    if (ebeln || belnr) {
                        return { 'Ord. de compra': ebeln, 'Documento no.': belnr };
                    }
                    return null;
                }).filter(item => item !== null);

                if (jsonData.length === 0) {
                     return reject(new Error("No se extrajeron datos de las columnas 'EBELN' y 'BELNR' del archivo de orden."));
                }

                resolve(jsonData as any[]);

            } catch (err: any) {
                reject(new Error(`Error al procesar el archivo HTML de orden ${file.name}: ${err.message}`));
            }
        };
        reader.onerror = () => reject(new Error(`Error al leer el archivo ${file.name}.`));
        reader.readAsText(file);
    });
};


export default function ReporteRetailPage() {
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [orderFile, setOrderFile] = useState<File | null>(null);
  const [processedData, setProcessedData] = useState<GroupedData[] | null>(null);
  const [status, setStatus] = useState<Status>('idle');
  const [progress, setProgress] = useState(0);
  const pdfContentRef = useRef<HTMLDivElement>(null);
  const { toast } = useToast();
  const dataFileInputRef = useRef<HTMLInputElement>(null);
  const orderFileInputRef = useRef<HTMLInputElement>(null);
  const cancelPdfGeneration = useRef(false);

  useEffect(() => {
    const generate = async () => {
      if (processedData && processedData.length > 0 && status === 'parsing') {
        await handleDownloadPdf();
      }
    };
    generate();
  }, [processedData, status]);

  const handleDataFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setDataFile(file);
    }
  };

  const handleOrderFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setOrderFile(file);
    }
  };

  const processData = (mainData: RetailData[], orderData: any[]): GroupedData[] | null => {
    // This helper normalizes PO numbers.
    const normalizePO = (value: any): string => {
        if (value === undefined || value === null) return "";
        return String(value).trim().replace(/^0+/, '');
    }

    // Accessor functions for specific fields.
    const getDocNo = (item: any) => getValue(item, ['documento no.', 'belnr']);
    const getPoNo = (item: any) => getValue(item, ['ord. de compra', 'ebeln', 'orden de compra']);

    const poToItemsMap = new Map<string, RetailData[]>();
    for (const item of mainData) {
        const po = getPoNo(item);
        if (po === undefined || po === null) continue;

        const poKey = normalizePO(po);
        if (!poKey) continue;
        
        if (!poToItemsMap.has(poKey)) {
            poToItemsMap.set(poKey, []);
        }
        poToItemsMap.get(poKey)!.push(item);
    }

    if (poToItemsMap.size === 0) {
        toast({ title: "Error en Archivo de Datos", description: "No se encontraron órdenes de compra válidas en el archivo de datos.", variant: "destructive" });
        return null;
    }

    const docToPoMap = new Map<string, Set<string>>();
    for (const item of orderData) {
        const docNo = getDocNo(item);
        const poNo = getPoNo(item);

        if (docNo === undefined || poNo === undefined) continue;

        const docKey = String(docNo).trim();
        const poKey = normalizePO(poNo);
        
        if (!docKey || !poKey) continue;

        if (!docToPoMap.has(docKey)) {
            docToPoMap.set(docKey, new Set());
        }
        docToPoMap.get(docKey)!.add(poKey);
    }
    
    if (docToPoMap.size === 0) {
        toast({ title: "Error en archivo de orden", description: "No se encontraron 'Documento no.' válidos en el archivo de orden.", variant: "destructive" });
        return null;
    }
    
    const finalGroupedData: GroupedData[] = [];
    
    for (const [docKey, poNumbers] of docToPoMap.entries()) {
        const allItemsForDoc: RetailData[] = [];
        for (const poKey of poNumbers) {
            const items = poToItemsMap.get(poKey);
            if (items) {
                allItemsForDoc.push(...items);
            }
        }

        if (allItemsForDoc.length > 0) {
            const totalCantidad = allItemsForDoc.reduce((sum, item) => sum + getQuantity(item), 0);
            const totalImporte = allItemsForDoc.reduce((sum, item) => sum + getAmount(item), 0);

            finalGroupedData.push({
                n_doc: docKey,
                items: allItemsForDoc,
                totalCantidad: totalCantidad,
                totalImporte: totalImporte,
            });
        }
    }

    if (finalGroupedData.length === 0) {
      toast({
        title: "No hay coincidencias",
        description: "No se encontraron datos en el archivo principal que correspondieran con las órdenes del archivo de orden.",
        variant: "destructive",
        duration: 9000
      });
      return null;
    }

    finalGroupedData.sort((a, b) => {
        const numA = Number(a.n_doc);
        const numB = Number(b.n_doc);
        if (!isNaN(numA) && !isNaN(numB)) {
            return numA - numB;
        }
        return String(a.n_doc).localeCompare(String(b.n_doc));
    });

    return finalGroupedData;
  };

  const handleGenerateReport = async () => {
    if (!dataFile || !orderFile) {
        toast({ title: "Faltan archivos", description: "Por favor, selecciona ambos archivos.", variant: "destructive" });
        return;
    }

    setStatus('parsing');
    
    try {
        const [mainData, orderData] = await Promise.all([
            readMainDataFile(dataFile),
            readOrderFileAsHTML(orderFile)
        ]);

        const groupedAndSortedData = processData(mainData, orderData);

        if (groupedAndSortedData && groupedAndSortedData.length > 0) {
            setProcessedData(groupedAndSortedData);
        } else {
            if (status !== 'error') {
              setStatus('error');
            }
        }

    } catch (error: any) {
        console.error("Error procesando archivos:", error);
        toast({
            title: "Error al procesar archivos",
            description: error.message || "Asegúrate de que los formatos sean correctos.",
            variant: "destructive",
            duration: 9000,
        });
        setStatus('error');
    }
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
        pdf.save('reporte_retail_ordenado.pdf');
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
    setDataFile(null);
    setOrderFile(null);
    setProcessedData(null);
    setStatus('idle');
    setProgress(0);
    cancelPdfGeneration.current = false;
    if(dataFileInputRef.current) dataFileInputRef.current.value = "";
    if(orderFileInputRef.current) orderFileInputRef.current.value = "";
  };

  const renderStatus = () => {
    switch (status) {
        case 'parsing':
            return (
                <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <Loader2 className="w-12 h-12 text-primary animate-spin" />
                    <p className="text-lg font-semibold text-foreground">Procesando tus archivos...</p>
                    <p className="text-sm text-muted-foreground">{dataFile?.name} & {orderFile?.name}</p>
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
                    <p className="text-sm text-muted-foreground">Tu descarga ha comenzado.</p>
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
                    <p className="text-sm text-muted-foreground">No se pudo generar el reporte. Por favor, revisa los archivos e inténtalo de nuevo.</p>
                    <Button variant="outline" onClick={resetState}>
                        <RefreshCw className="mr-2 h-4 w-4" />
                        Intentar de Nuevo
                    </Button>
                </div>
            );
        case 'idle':
        default:
            return (
              <div className="flex flex-col items-center justify-center space-y-6 text-center w-full">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 w-full">
                      <div className="flex flex-col items-center p-6 border-2 border-dashed rounded-lg">
                          <UploadCloud className="w-12 h-12 text-primary/70" />
                          <p className="font-semibold mt-2">1. Archivo de Datos</p>
                          <p className="text-xs text-muted-foreground">El reporte principal de SAP.</p>
                          <Button onClick={() => dataFileInputRef.current?.click()} className="mt-4" variant="outline">
                              Seleccionar Archivo
                          </Button>
                          {dataFile && <p className="text-sm mt-2 text-green-600 font-medium">{dataFile.name}</p>}
                      </div>
                      <div className="flex flex-col items-center p-6 border-2 border-dashed rounded-lg">
                          <UploadCloud className="w-12 h-12 text-primary/70" />
                          <p className="font-semibold mt-2">2. Archivo de Orden</p>
                          <p className="text-xs text-muted-foreground">HTML de SAP guardado como .xls.</p>
                          <Button onClick={() => orderFileInputRef.current?.click()} className="mt-4" variant="outline">
                              Seleccionar Archivo
                          </Button>
                          {orderFile && <p className="text-sm mt-2 text-green-600 font-medium">{orderFile.name}</p>}
                      </div>
                  </div>

                  <Button onClick={handleGenerateReport} disabled={!dataFile || !orderFile} size="lg" className="w-full md:w-auto">
                      <Shuffle className="mr-2 h-4 w-4" />
                      Generar Reporte PDF
                  </Button>

                  <input ref={dataFileInputRef} type="file" className="hidden" onChange={handleDataFileChange} accept=".xlsx, .xls" />
                  <input ref={orderFileInputRef} type="file" className="hidden" onChange={handleOrderFileChange} accept=".xlsx, .xls" />
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
          Reportes de Retail (con Ordenamiento)
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube tu reporte de SAP y un archivo de orden para generar un PDF listo para imprimir, agrupado y ordenado por número de documento.
        </p>
      </div>

      <Card className="max-w-4xl mx-auto shadow-lg">
          <CardContent className="p-8 min-h-[350px] flex items-center justify-center">
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
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Descripcion Material</TableHead>
                      <TableHead className="px-1 py-1 text-black font-bold border border-neutral-300">Fecha Ingreso</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Cant. In.</TableHead>
                      <TableHead className="text-right px-1 py-1 text-black font-bold border border-neutral-300">Costo Total</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                  {group.items.map((item, itemIndex) => {
                      let formattedDate = '';
                      const dateValue = getPostingDate(item);
                      if (dateValue) {
                        try {
                           // Handle Excel's numeric date format
                           if (typeof dateValue === 'number' && dateValue > 1) {
                                const excelEpoch = new Date(1899, 11, 30);
                                const date = new Date(excelEpoch.getTime() + dateValue * 24 * 60 * 60 * 1000);
                                formattedDate = `${String(date.getDate()).padStart(2, '0')}.${String(date.getMonth() + 1).padStart(2, '0')}.${date.getFullYear()}`;
                           } else { // Handle string dates
                                const date = new Date(dateValue);
                                if (!Number.isNaN(date.getTime())) {
                                    const adjustedDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60 * 1000);
                                    formattedDate = `${String(adjustedDate.getDate()).padStart(2, '0')}.${String(adjustedDate.getMonth() + 1).padStart(2, '0')}.${adjustedDate.getFullYear()}`;
                                }
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

    