"use client";

import { useState, useRef, ChangeEvent } from 'react';
import { ArrowLeft, Shuffle, FileSpreadsheet, Loader2, FileCheck2 } from 'lucide-react';
import Link from 'next/link';
import * as XLSX from 'xlsx';

import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Separator } from '@/components/ui/separator';
import { cn } from '@/lib/utils';

export default function ReporteRetailPage() {
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [processCompleted, setProcessCompleted] = useState(false);
  const [fileName, setFileName] = useState<{data: string | null, excel: string | null}>({ data: null, excel: null });
  const [debugData, setDebugData] = useState<{rawText: string, orders: string[]} | null>(null);
  const dataInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>, type: 'data' | 'excel') => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(file.type)) {
      toast({ title: "Error de archivo", description: "Por favor, sube un archivo de Excel (.xlsx o .xls).", variant: "destructive" });
      return;
    }
    
    if (type === 'data') setDataFile(file);
    if (type === 'excel') setExcelFile(file);
    setFileName(prev => ({ ...prev, [type]: file.name }));
    setDebugData(null);
    setProcessCompleted(false);
  };

  const handleProcess = async () => {
    if (!dataFile || !excelFile) {
      toast({ title: "Faltan archivos", description: "Por favor, sube ambos archivos de Excel.", variant: "destructive" });
      return;
    }

    setLoading(true);
    setDebugData(null);
    setProcessCompleted(false);
    
    try {
        // 1. Read and parse ordering Excel (EKBE)
        const orderExcelBuffer = await excelFile.arrayBuffer();
        const orderWorkbook = XLSX.read(orderExcelBuffer, { type: 'buffer' });
        const orderSheetName = orderWorkbook.SheetNames[0];
        const orderWorksheet = orderWorkbook.Sheets[orderSheetName];
        
        let orderHeaderRowIndex = -1;
        const orderRange = XLSX.utils.decode_range(orderWorksheet['!ref']);
        for(let R = orderRange.s.r; R <= Math.min(orderRange.e.r, 10); ++R) {
            const rowValues = [];
            for (let C = orderRange.s.c; C <= orderRange.e.c; ++C) {
                const cell = orderWorksheet[XLSX.utils.encode_cell({r:R, c:C})];
                if (cell && cell.v) {
                    rowValues.push(String(cell.v).trim().toLowerCase());
                }
            }
            if (rowValues.includes('ebeln') && rowValues.includes('belnr')) {
                orderHeaderRowIndex = R;
                break;
            }
        }

        if (orderHeaderRowIndex === -1) {
             throw new Error("No se pudo encontrar la fila de encabezado con 'EBELN' y 'BELNR' en el archivo Excel de ordenamiento. Asegúrate de que el archivo exportado de SAP contenga estas columnas.");
        }
        
        const orderJson = XLSX.utils.sheet_to_json(orderWorksheet, { range: orderHeaderRowIndex }) as any[];

        if (orderJson.length === 0) {
            throw new Error("El archivo Excel de ordenamiento está vacío o no tiene el formato esperado.");
        }
        
        const belnrHeader = Object.keys(orderJson[0]).find(h => h.trim().toLowerCase() === 'belnr');
        const ebelnHeader = Object.keys(orderJson[0]).find(h => h.trim().toLowerCase() === 'ebeln');

        if (!belnrHeader || !ebelnHeader) {
            throw new Error(`El archivo Excel de ordenamiento debe contener las columnas 'BELNR' y 'EBELN'.`);
        }

        const uniquePairs = new Map<string, { belnr: string, ebeln: string }>();
        orderJson.forEach((row: any) => {
            const belnr = String(row[belnrHeader]).trim();
            const ebeln = String(row[ebelnHeader]).trim();
            if (belnr && ebeln) {
                const key = `${belnr}-${ebeln}`;
                if (!uniquePairs.has(key)) {
                    uniquePairs.set(key, { belnr, ebeln });
                }
            }
        });

        let processedOrderData = Array.from(uniquePairs.values());
        processedOrderData.sort((a, b) => {
            const belnrCompare = a.belnr.localeCompare(b.belnr, undefined, { numeric: true });
            if (belnrCompare !== 0) return belnrCompare;
            return a.ebeln.localeCompare(b.ebeln, undefined, { numeric: true });
        });
        
        const orderedEbelns = processedOrderData.map(p => p.ebeln);
        const uniqueOrderedEbelns = [...new Set(orderedEbelns)];
        
        if (uniqueOrderedEbelns.length === 0) {
            throw new Error("No se encontraron valores de 'EBELN' y 'BELNR' válidos en el archivo de ordenamiento.");
        }
        
        // 2. Read data Excel
        const dataExcelBuffer = await dataFile.arrayBuffer();
        const dataWorkbook = XLSX.read(dataExcelBuffer, { type: 'buffer' });
        const dataSheetName = dataWorkbook.SheetNames[0];
        const dataWorksheet = dataWorkbook.Sheets[dataSheetName];
        const dataJson = XLSX.utils.sheet_to_json(dataWorksheet) as any[];

        if (dataJson.length === 0) {
            throw new Error("El archivo de datos Excel está vacío.");
        }

        let ordCompraHeader = Object.keys(dataJson[0]).find(h => h.trim().toLowerCase().replace(/\./g, '') === 'ord de compra');
         if (!ordCompraHeader) {
            ordCompraHeader = Object.keys(dataJson[0]).find(h => h.trim().toLowerCase() === 'ebeln');
        }

        if (!ordCompraHeader) {
            const foundHeaders = Object.keys(dataJson[0]).join(', ');
             setDebugData({
                rawText: `Columnas encontradas en el Excel de datos: ${foundHeaders}`,
                orders: ["Se esperaba el encabezado: 'Ord. de Compra' o 'EBELN'"],
             })
            throw new Error(`El archivo de datos debe contener la columna 'Ord. de Compra' o 'EBELN'.`);
        }

        // 3. Group data rows by 'Ord. de Compra'
        const dataByEbeln = new Map<string, any[]>();
        dataJson.forEach(row => {
            const ebeln = String((row as any)[ordCompraHeader!]).trim();
            if (!ebeln) return;
            if (!dataByEbeln.has(ebeln)) {
                dataByEbeln.set(ebeln, []);
            }
            dataByEbeln.get(ebeln)!.push(row);
        });

        // 4. Create new sorted data array
        const sortedData: any[] = [];
        const usedEbelns = new Set<string>();

        uniqueOrderedEbelns.forEach(ebeln => {
            if (dataByEbeln.has(ebeln)) {
                sortedData.push(...dataByEbeln.get(ebeln)!);
                usedEbelns.add(ebeln);
            }
        });
        
        let unmappedRowsCount = 0;
        dataJson.forEach(row => {
            const ebeln = String((row as any)[ordCompraHeader!]).trim();
            if (ebeln && !usedEbelns.has(ebeln)) {
                if(dataByEbeln.has(ebeln)) {
                    const rowsToAdd = dataByEbeln.get(ebeln)!;
                    sortedData.push(...rowsToAdd);
                    usedEbelns.add(ebeln);
                    unmappedRowsCount += rowsToAdd.length;
                }
            }
        });

        if (sortedData.length === 0) {
            const ebelnsForDebug = uniqueOrderedEbelns.slice(0, 20).join(', ');
            setDebugData({
                 rawText: `Cabeceras del archivo de datos: ${Object.keys(dataJson[0]).join(', ')}`,
                 orders: [`Valores 'Ord. de Compra' / 'EBELN' buscados (primeros 20): ${ebelnsForDebug}`],
             });
            throw new Error("No se pudo hacer coincidir ningún dato entre los dos archivos. Revisa que los 'Ord. de Compra' existan en ambos archivos.");
        }

        // 5. Create and download new Excel file
        const newWorksheet = XLSX.utils.json_to_sheet(sortedData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Reporte Ordenado');
        
        const newExcelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([newExcelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        handleDownload(url, `ordenado_${dataFile.name}`);

        toast({
          title: "Proceso completado",
          description: `Se reordenaron las filas. ${unmappedRowsCount > 0 ? `${unmappedRowsCount} filas no mapeadas se añadieron al final.` : ''}`,
        });
        setProcessCompleted(true);

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
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube dos archivos de Excel para generar un nuevo reporte ordenado.
        </p>
      </div>

      <Card className="max-w-4xl mx-auto shadow-lg">
        <CardHeader>
          <CardTitle>Organizador de Reportes Excel</CardTitle>
          <CardDescription>Sube los dos archivos de Excel requeridos para comenzar el proceso.</CardDescription>
        </CardHeader>
        <CardContent className="p-6 grid md:grid-cols-2 gap-8">
          <div className={cn(
            "flex flex-col items-center justify-center space-y-4 p-6 border-2 border-dashed rounded-lg transition-colors",
            dataFile ? "border-green-500 bg-green-50 dark:bg-green-500/10" : "border-input"
          )}>
            {dataFile ? (
              <FileCheck2 className="w-12 h-12 text-green-500" />
            ) : (
              <FileSpreadsheet className="w-12 h-12 text-primary" />
            )}
            <p className="text-lg font-semibold text-foreground text-center">
              {dataFile ? "Datos cargados" : "1. Sube el Excel con los datos"}
            </p>
            <p className="text-sm text-muted-foreground h-5 text-center truncate max-w-full px-2" title={fileName.data ?? ''}>
                {fileName.data}
            </p>
            <Button variant="outline" onClick={() => dataInputRef.current?.click()}>
              {dataFile ? "Cambiar Excel" : "Seleccionar Excel"}
            </Button>
            <input ref={dataInputRef} type="file" className="hidden" onChange={(e) => handleFileChange(e, 'data')} accept=".xlsx, .xls" />
          </div>

          <div className={cn(
            "flex flex-col items-center justify-center space-y-4 p-6 border-2 border-dashed rounded-lg transition-colors",
            excelFile ? "border-green-500 bg-green-50 dark:bg-green-500/10" : "border-input"
          )}>
            {excelFile ? (
              <FileCheck2 className="w-12 h-12 text-green-500" />
            ) : (
              <FileSpreadsheet className="w-12 h-12 text-primary" />
            )}
            <p className="text-lg font-semibold text-foreground text-center">
              {excelFile ? "Orden cargado" : "2. Sube el Excel con el orden"}
            </p>
            <p className="text-sm text-muted-foreground h-5 text-center truncate max-w-full px-2" title={fileName.excel ?? ''}>
                {fileName.excel}
            </p>
            <Button variant="outline" onClick={() => excelInputRef.current?.click()}>
              {excelFile ? "Cambiar Excel" : "Seleccionar Excel"}
            </Button>
            <input ref={excelInputRef} type="file" className="hidden" onChange={(e) => handleFileChange(e, 'excel')} accept=".xlsx, .xls" />
          </div>
        </CardContent>
        <div className="p-6 pt-0 flex justify-center">
            {loading ? (
                 <div className="flex flex-col items-center justify-center space-y-4 text-center">
                    <Loader2 className="w-12 h-12 text-primary animate-spin" />
                    <p className="text-lg font-semibold text-foreground">Procesando y reordenando el Excel...</p>
                    <p className="text-sm text-muted-foreground">Esto puede tomar un momento.</p>
                </div>
            ) : (
                <Button 
                    size="lg" 
                    onClick={handleProcess} 
                    disabled={!dataFile || !excelFile}
                    className={cn(processCompleted && "bg-green-600 hover:bg-green-700")}
                >
                  {processCompleted ? <FileCheck2 className="mr-2 h-5 w-5" /> : <Shuffle className="mr-2 h-5 w-5" />}
                  {processCompleted ? '¡Ordenado! Descargar de nuevo' : 'Ordenar Excel'}
                </Button>
            )}
        </div>
      </Card>

      {debugData && (
        <Card className="max-w-4xl mx-auto mt-8">
          <CardHeader>
            <CardTitle>Información de Depuración</CardTitle>
            <CardDescription>
              No se pudo ordenar ninguna fila. Revisa que la columna 'Ord. de Compra' del Excel de datos coincida con los valores 'EBELN' del Excel de orden.
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div>
              <h4 className="font-semibold mb-2">Información del archivo de datos:</h4>
              <pre className="p-4 bg-muted rounded-md text-xs whitespace-pre-wrap max-h-[300px] overflow-auto">
                {debugData.rawText}
              </pre>
            </div>
            <div>
              <h4 className="font-semibold mb-2">Información del archivo de ordenamiento:</h4>
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
                  <h4 className="font-bold text-lg mb-3 text-foreground/90">Parte 1: Generar el Reporte de Datos (Excel)</h4>
                  <ol className="list-decimal list-inside space-y-4 text-foreground/80">
                    <li>
                      <strong>Generar el reporte principal desde SAP</strong>. Este archivo debe contener toda la información detallada que deseas ordenar.
                    </li>
                    <li>
                      Asegúrate de que el archivo contenga una columna llamada <strong>'Ord. de Compra'</strong> o <strong>'EBELN'</strong>, ya que se usará para el ordenamiento.
                    </li>
                     <li>
                      Exporta este reporte como un archivo de <strong>Excel</strong>.
                    </li>
                  </ol>
                </div>
                <Separator />
                <div>
                  <h4 className="font-bold text-lg mb-3 text-foreground/90">Parte 2: Generar el Archivo de Ordenamiento (Excel)</h4>
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
                        <li>En el campo <strong>`BELNR`</strong>, usa la selección múltiple para pegar todos los números de documento que correspondan a tu reporte principal.</li>
                        <li>Ejecuta la selección (F8).</li>
                      </ul>
                    </li>
                    <li>
                      <strong>Exportar a Excel</strong>:
                      <ul className="list-disc list-inside pl-5 mt-2 space-y-1">
                        <li>Asegúrate de que las columnas `BELNR` y `EBELN` estén visibles.</li>
                        <li>Ve a <strong>Sistema &gt; Lista &gt; Grabar &gt; Fichero local</strong>.</li>
                        <li>Elige la opción <strong>"Hoja de cálculo"</strong>.</li>
                        <li>Guarda el archivo. Este será el archivo de ordenamiento que subirás a la herramienta.</li>
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
              Selecciona el Excel con los datos y el Excel con el orden. El de orden debe tener las columnas 'BELNR' y 'EBELN'.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">2</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">La Herramienta Procesa</h4>
            <p className="text-foreground/80">
              La aplicación lee tu Excel de ordenamiento y lo usa para reordenar las filas de tu Excel de datos según la columna 'Ord. de Compra'.
            </p>
          </div>
          <div className="flex flex-col items-center space-y-2">
            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-primary/10 text-primary mb-4">
              <span className="text-2xl font-bold">3</span>
            </div>
            <h4 className="text-xl font-semibold text-foreground">Descarga el Excel Ordenado</h4>
            <p className="text-foreground/80">
              Se genera un nuevo archivo Excel con todas las filas reordenadas secuencialmente según el orden de tu archivo de ordenamiento.
            </p>
          </div>
        </div>
      </section>
    </main>
  );
}

    