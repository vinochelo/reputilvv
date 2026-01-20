"use client";

import { useState, useRef, ChangeEvent } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { UploadCloud, FileDown, Loader2, FileX2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow, TableFooter } from '@/components/ui/table';
import { useToast } from '@/hooks/use-toast';
import type { ExcelData, GroupedData } from '@/lib/types';
import { mockExcelData } from '@/lib/mock-data';

export default function Home() {
  const [processedData, setProcessedData] = useState<GroupedData[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const reportRefs = useRef<(HTMLDivElement | null)[]>([]);
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

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
      
      // Simulating Excel parsing. In a real app, you would use a library like 'xlsx' here.
      // Using mock data for this demonstration.
      processData(mockExcelData);
    }
  };

  const processData = (data: ExcelData[]) => {
    const grouped = data.reduce<GroupedData[]>((acc, item) => {
      let group = acc.find(g => g.n_doc === item['N doc']);
      if (!group) {
        group = {
          n_doc: item['N doc'],
          cliente: item['Nombre cliente'],
          rut: `${item['Rut']}-${item['Dv']}`,
          items: [],
          totalPagar: 0,
        };
        acc.push(group);
      }
      group.items.push(item);
      group.totalPagar += item['Valor a pagar'];
      return acc;
    }, []);
    setProcessedData(grouped);
    reportRefs.current = grouped.map(() => null);
  };

  const handleDownloadPdf = async () => {
    if (!processedData || reportRefs.current.length === 0) return;

    setLoading(true);
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    
    for (let i = 0; i < processedData.length; i++) {
      const reportElement = reportRefs.current[i];
      if (reportElement) {
        if (i > 0) {
          pdf.addPage();
        }
        const canvas = await html2canvas(reportElement, { scale: 2 });
        const imgData = canvas.toDataURL('image/png');
        const imgProps = pdf.getImageProperties(imgData);
        const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
        
        pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);
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
          </CardContent>
        </Card>
      ) : (
        <div className="space-y-8">
          <Card className="max-w-4xl mx-auto shadow-lg">
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
                <div id="pdf-content" className="space-y-8">
                {processedData.map((group, index) => (
                    <div key={group.n_doc} ref={el => reportRefs.current[index] = el} className="p-6 bg-white text-black border rounded-lg shadow-sm">
                      <header className="mb-4">
                          <h2 className="font-headline text-xl font-bold">Reporte de Documento</h2>
                          <div className="grid grid-cols-2 text-sm mt-2 gap-x-4">
                              <p><span className="font-semibold">N° Doc:</span> {group.n_doc}</p>
                              <p><span className="font-semibold">Cliente:</span> {group.cliente}</p>
                              <p><span className="font-semibold">RUT:</span> {group.rut}</p>
                          </div>
                      </header>
                      <Table>
                          <TableHeader>
                          <TableRow>
                              <TableHead className="font-headline">Tipo doc</TableHead>
                              <TableHead className="font-headline">Folio</TableHead>
                              <TableHead className="font-headline">Fecha Emisión</TableHead>
                              <TableHead className="font-headline">Fecha Venc.</TableHead>
                              <TableHead className="font-headline text-right">Valor a pagar</TableHead>
                          </TableRow>
                          </TableHeader>
                          <TableBody>
                          {group.items.map((item, itemIndex) => (
                              <TableRow key={itemIndex}>
                              <TableCell>{item['Tipo doc']}</TableCell>
                              <TableCell>{item.Folio}</TableCell>
                              <TableCell>{item['Fecha emision']}</TableCell>
                              <TableCell>{item['Fecha venc']}</TableCell>
                              <TableCell className="text-right">{item['Valor a pagar'].toLocaleString('es-CL', { style: 'currency', currency: 'CLP' })}</TableCell>
                              </TableRow>
                          ))}
                          </TableBody>
                          <TableFooter>
                          <TableRow className="bg-accent/30 hover:bg-accent/40">
                              <TableCell colSpan={4} className="font-headline text-right font-bold text-lg">Total</TableCell>
                              <TableCell className="font-headline text-right font-bold text-lg">
                              {group.totalPagar.toLocaleString('es-CL', { style: 'currency', currency: 'CLP' })}
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
