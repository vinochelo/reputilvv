"use client";

import { useState, useRef, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileDown, Loader2, FileX2, ArrowLeft } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import Link from 'next/link';
import { Textarea } from '@/components/ui/textarea';

export default function ControlRetencionesPage() {
  const [processedContent, setProcessedContent] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (file.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' && file.type !== 'application/vnd.ms-excel') {
      toast({
        title: "Error de archivo",
        description: "Por favor, sube un archivo de Excel (.xlsx o .xls).",
        variant: "destructive",
      });
      return;
    }

    setFileName(file.name);
    setLoading(true);
    setProcessedContent(null);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<any>(worksheet);

        if (jsonData.length === 0) {
            toast({
                title: "Archivo vacío",
                description: "El archivo de Excel no contiene datos.",
                variant: "destructive",
            });
            resetState();
            return;
        }
        
        const requiredColumns = ['RUC', 'COMPROBANTE', 'COD. COMP.', 'Nro. Autorizacion', 'BASE IMPONIBLE', 'FECHA EMISION'];
        const firstRow = jsonData[0];
        const missingColumns = requiredColumns.filter(col => !(col in firstRow));

        if (missingColumns.length > 0) {
            toast({
                title: "Columnas faltantes",
                description: `El archivo no tiene las siguientes columnas requeridas: ${missingColumns.join(', ')}`,
                variant: "destructive",
            });
            resetState();
            return;
        }

        let processed = '';
        jsonData.forEach((row) => {
          const ruc = String(row.RUC).padStart(13, '0');
          
          // Replicating the logic from the original repository
          const comprobante = String(row.COMPROBANTE).padStart(9, '0');
          const serie = comprobante.slice(0, 7);

          const tipoComp = String(row['COD. COMP.']).padStart(2, '0');
          const autorizacion = row['Nro. Autorizacion'];
          const baseImponible = parseFloat(row['BASE IMPONIBLE']).toFixed(2);
          
          const excelDate = row['FECHA EMISION'];
          const jsDate = new Date((excelDate - (25567 + 2)) * 86400 * 1000);
          
          const day = String(jsDate.getDate()).padStart(2, '0');
          const month = String(jsDate.getMonth() + 1).padStart(2, '0');
          const year = jsDate.getFullYear();
          const formattedDate = `${day}${month}${year}`;
          
          processed += `01|${formattedDate}|03|${ruc}|${tipoComp}|${serie}|${comprobante}|${autorizacion}|${baseImponible}|0|0|0\r\n`;
        });

        if (!processed) {
            toast({
                title: "No se procesaron datos",
                description: "No se encontraron filas válidas en el archivo.",
                variant: "destructive",
            });
            resetState();
            return;
        }

        setProcessedContent(processed);

      } catch (error) {
        console.error("Error processing Excel file:", error);
        toast({
          title: "Error al procesar el archivo",
          description: "No se pudo leer el archivo. Asegúrate de que el formato y las columnas sean correctas.",
          variant: "destructive",
        });
        resetState();
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => {
      console.error("Error reading file");
      toast({
          title: "Error al leer el archivo",
          description: "Ocurrió un error al intentar leer el archivo.",
          variant: "destructive",
      });
      resetState();
    }
    reader.readAsBinaryString(file);
  };
  
  const handleDownload = () => {
    if (!processedContent || !fileName) return;
    const blob = new Blob([processedContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const txtFileName = fileName.replace(/\.[^/.]+$/, "") + ".txt";
    link.download = txtFileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    toast({
      title: "Descarga iniciada",
      description: `Se ha descargado el archivo ${txtFileName}.`
    });
  };

  const resetState = () => {
    setProcessedContent(null);
    setFileName(null);
    setLoading(false);
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
          Extractor de Retenciones
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube un archivo de Excel para generar un archivo de texto (.txt) compatible con la importación en DIMM.
        </p>
      </div>

      {!processedContent ? (
        <Card className="max-w-xl mx-auto shadow-lg border-2 border-dashed border-primary/50 hover:border-primary transition-colors">
          <CardContent className="p-8">
            {loading ? (
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
        <Card className="max-w-4xl mx-auto shadow-lg">
          <CardHeader className="flex flex-row items-center justify-between">
              <div>
                  <CardTitle className="font-headline text-2xl">Previsualización y Descarga</CardTitle>
                  <CardDescription>Archivo: {fileName}</CardDescription>
              </div>
              <div className="flex gap-2">
                  <Button variant="outline" onClick={resetState}>
                    <FileX2 className="mr-2 h-4 w-4" />
                    Cargar Otro
                  </Button>
                  <Button onClick={handleDownload}>
                    <FileDown className="mr-2 h-4 w-4" />
                    Descargar .txt
                  </Button>
              </div>
          </CardHeader>
          <CardContent>
              <Textarea
                readOnly
                value={processedContent}
                className="h-64 font-mono text-xs bg-muted/50"
                placeholder="Contenido del archivo generado..."
              />
          </CardContent>
        </Card>
      )}
    </main>
  );
}
