"use client";

import { useState, useRef, ChangeEvent } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, FileX2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow, TableFooter } from '@/components/ui/table';
import { useToast } from '@/hooks/use-toast';
import type { ExcelData, GroupedData } from '@/lib/types';

export default function Home() {
  const [processedData, setProcessedData] = useState<GroupedData[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const pdfContentRef = useRef<HTMLDivElement>(null);
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const getDocId = (item: any) => {
    // Handles variations like 'Nº doc.', 'N° doc.', 'Nro. Doc.', etc.
    const keys = Object.keys(item);
    const docIdKey = keys.find(key => key.toLowerCase().trim().replace('.', '') === 'nº doc' || key.toLowerCase().trim().replace('.', '') === 'n° doc');
    return docIdKey ? item[docIdKey] : undefined;
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
      
      group.items.push(item);
      group.totalCantidad += Number(item['Cantidad'] ?? 0);
      group.totalCostoTotal += Number(item['Costo Total'] ?? 0);
      group.totalPrecioVenta += Number(item['Precio Venta'] ?? 0);
      group.totalValorAPagar += Number(item['Valor a pagar'] ?? 0);
      group.totalUtilidad += Number(item['Utilidad %'] ?? 0);
      return acc;
    }, []);

    if (grouped.length === 0) {
      toast({
        title: "No se pudo agrupar",
        description: "No se encontraron datos para agrupar. Revisa que la columna de 'Nº doc.' exista y tenga valores.",
        variant: "destructive",
      });
      return;
    }

    setProcessedData(grouped);
  };

  const handleDownloadPdf = async () => {
    if (!pdfContentRef.current) return;

    setLoading(true);
    await new Promise(resolve => setTimeout(resolve, 0)); // Allow UI to update
    
    const pdf = new jsPDF('l', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const reportElements = Array.from(pdfContentRef.current.children);
    
    for (let i = 0; i < reportElements.length; i++) {
      const reportElement = reportElements[i] as HTMLElement;
      if (reportElement) {
        if (i > 0) {
          pdf.addPage();
        }
        try {
            const canvas = await html2canvas(reportElement, { scale: 2 });
            const imgData = canvas.toDataURL('image/png');
            const imgProps = pdf.getImageProperties(imgData);
            const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
            pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);
        } catch(e) {
            console.error(e);
            toast({
              title: "Error al generar PDF",
              description: "No se pudo procesar uno de los reportes para el PDF.",
              variant: "destructive",
            });
        }
      }
    }

    pdf.save('reporte.pdf');
    setLoading(false);
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
      <div className="text-center mb-12">
        <h1 className="text-4xl font-headline font-bold tracking-tight text-primary-foreground sm:text-5xl lg:text-6xl">
          Conversor de Excel a PDF
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube tu archivo de Excel para generar un reporte en PDF listo para imprimir.
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
                    <Button variant="outline" onClick={resetState}>
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
                <div id="pdf-content" ref={pdfContentRef} className="space-y-8">
                {processedData.map((group, index) => (
                    <div key={`${group.n_doc}-${index}`} className="p-6 bg-white text-black border rounded-lg shadow-sm">
                      <header className="mb-4">
                          <h2 className="font-headline text-xl font-bold">Reporte utilidad venta en verde</h2>
                          <div className="text-sm mt-2">
                              <p><span className="font-semibold">N° Doc:</span> {group.n_doc}</p>
                          </div>
                      </header>
                      <Table>
                          <TableHeader>
                            <TableRow>
                              <TableHead className="font-headline">Documento material</TableHead>
                              <TableHead className="font-headline">Factura</TableHead>
                              <TableHead className="font-headline">Nº doc.</TableHead>
                              <TableHead className="font-headline">Centro</TableHead>
                              <TableHead className="font-headline">Fecha Factura</TableHead>
                              <TableHead className="font-headline">Proveedor</TableHead>
                              <TableHead className="font-headline">Nombre del proveedor</TableHead>
                              <TableHead className="font-headline">Material</TableHead>
                              <TableHead className="font-headline">Texto breve de material</TableHead>
                              <TableHead className="font-headline text-right">Cantidad</TableHead>
                              <TableHead className="font-headline text-right">Costo Total</TableHead>
                              <TableHead className="font-headline text-right">Precio Venta</TableHead>
                              <TableHead className="font-headline text-right">Utilidad %</TableHead>
                              <TableHead className="font-headline text-right">Valor a pagar</TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                          {group.items.map((item, itemIndex) => (
                              <TableRow key={itemIndex}>
                                <TableCell>{item['Documento material']}</TableCell>
                                <TableCell>{item['Factura']}</TableCell>
                                <TableCell>{getDocId(item)}</TableCell>
                                <TableCell>{item['Centro']}</TableCell>
                                <TableCell>{new Date(item['Fecha Factura']).toLocaleDateString('es-CL')}</TableCell>
                                <TableCell>{item['Proveedor']}</TableCell>
                                <TableCell>{item['Nombre del proveedor']}</TableCell>
                                <TableCell>{item['Material']}</TableCell>
                                <TableCell>{item['Texto breve de material']}</TableCell>
                                <TableCell className="text-right">{(item['Cantidad'] ?? 0).toFixed(3)}</TableCell>
                                <TableCell className="text-right">{(item['Costo Total'] ?? 0).toFixed(2)}</TableCell>
                                <TableCell className="text-right">{(item['Precio Venta'] ?? 0).toFixed(2)}</TableCell>
                                <TableCell className="text-right">{(item['Utilidad %'] ?? 0).toFixed(2)}</TableCell>
                                <TableCell className="text-right">{(item['Valor a pagar'] ?? 0).toFixed(2)}</TableCell>
                              </TableRow>
                          ))}
                          </TableBody>
                          <TableFooter>
                            <TableRow className="bg-accent/30 hover:bg-accent/40">
                              <TableCell colSpan={9} className="font-headline text-right font-bold text-lg">Total</TableCell>
                              <TableCell className="font-headline text-right font-bold text-lg">
                                {group.totalCantidad.toFixed(3)}
                              </TableCell>
                              <TableCell className="font-headline text-right font-bold text-lg">
                                {group.totalCostoTotal.toFixed(2)}
                              </TableCell>
                              <TableCell className="font-headline text-right font-bold text-lg">
                                {group.totalPrecioVenta.toFixed(2)}
                              </TableCell>
                              <TableCell className="font-headline text-right font-bold text-lg">
                                {group.items.length > 0 ? (group.totalUtilidad / group.items.length).toFixed(2) : '0.00'}
                              </TableCell>
                              <TableCell className="font-headline text-right font-bold text-lg">
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
    </main>
  );
}
