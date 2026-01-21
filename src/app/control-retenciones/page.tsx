"use client";

import { useState, useRef, ChangeEvent } from 'react';
import { UploadCloud, FileDown, Loader2, FileX2, ArrowLeft, BrainCircuit } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import Link from 'next/link';
import { Textarea } from '@/components/ui/textarea';
import { extractRetenciones } from '@/ai/flows/extract-retenciones-flow';

export default function ControlRetencionesPage() {
  const [processedContent, setProcessedContent] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (file.type !== 'application/pdf') {
      toast({
        title: "Error de archivo",
        description: "Por favor, sube un archivo PDF.",
        variant: "destructive",
      });
      return;
    }

    setFileName(file.name);
    setLoading(true);
    setProcessedContent(null);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const dataUri = e.target?.result as string;
        if (!dataUri) {
          throw new Error("No se pudo leer el archivo.");
        }
        
        const result = await extractRetenciones(dataUri);
        setProcessedContent(result);

      } catch (error: any) {
        console.error("Error processing PDF file:", error);
        toast({
          title: "Error al procesar el archivo con IA",
          description: error.message || "No se pudo extraer la información del PDF. Intenta con otro archivo.",
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
    reader.readAsDataURL(file);
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
          Extractor de Retenciones (PDF)
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Sube un archivo de retenciones en PDF para que la IA extraiga los datos y genere un archivo .txt compatible con DIMM.
        </p>
      </div>

      {!processedContent ? (
        <Card className="max-w-xl mx-auto shadow-lg border-2 border-dashed border-primary/50 hover:border-primary transition-colors">
          <CardContent className="p-8">
            {loading ? (
              <div className="flex flex-col items-center justify-center space-y-4 text-center">
                  <Loader2 className="w-16 h-16 text-primary animate-spin" />
                  <p className="text-lg font-semibold text-foreground">La IA está procesando tu PDF...</p>
                  <p className="text-sm text-muted-foreground">Esto puede tardar unos segundos.</p>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center space-y-4 text-center">
                <UploadCloud className="w-16 h-16 text-primary" />
                <p className="text-lg font-semibold text-foreground">Arrastra y suelta tu archivo PDF aquí</p>
                <p className="text-muted-foreground">o</p>
                <Button onClick={() => fileInputRef.current?.click()}>
                  <BrainCircuit className="mr-2 h-4 w-4" />
                  Seleccionar PDF
                </Button>
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  onChange={handleFileChange}
                  accept="application/pdf"
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
                    Procesar Otro
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
