"use client";

import { useState, useRef, ChangeEvent } from 'react';
import { UploadCloud, FileDown, Loader2, ArrowLeft, BrainCircuit, Mail, Send, Copy, ExternalLink, Trash2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import Link from 'next/link';
import { extractRetenciones } from '@/ai/flows/extract-retenciones-flow';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Checkbox } from '@/components/ui/checkbox';
import { Badge } from '@/components/ui/badge';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';

type Retention = {
  id: string;
  nroRetencion: string;
  razonSocial: string;
  nroFactura: string;
  valor: string;
  estado: 'Solicitado' | 'Anulado';
  fechaCreacion: string;
  fechaEmision: string;
  autorizacion: string;
};

const initialRetentions: Retention[] = [
  {
    id: 'ret1',
    nroRetencion: '005-001-000210535',
    razonSocial: 'CORPMUNAB SOCIEDAD ANONIMA',
    nroFactura: '002200000056785',
    valor: '0.45',
    estado: 'Solicitado',
    fechaCreacion: '21/01/2026 10:22',
    fechaEmision: '20/01/2026',
    autorizacion: '200120260',
  },
  {
    id: 'ret2',
    nroRetencion: '005-001-000210452',
    razonSocial: 'HARDOOMSOLUTIONS S.A.',
    nroFactura: '001002000006218',
    valor: '40.94',
    estado: 'Solicitado',
    fechaCreacion: '20/01/2026 07:20',
    fechaEmision: '16/01/2026',
    autorizacion: '160120260',
  },
  {
    id: 'ret3',
    nroRetencion: '005-001-000210451',
    razonSocial: 'HARDOOMSOLUTIONS S.A.',
    nroFactura: '001002000006224',
    valor: '40.94',
    estado: 'Solicitado',
    fechaCreacion: '20/01/2026 07:19',
    fechaEmision: '16/01/2026',
    autorizacion: '160120260',
  },
];


export default function ControlRetencionesBetaPage() {
  const [retentions, setRetentions] = useState<Retention[]>(initialRetentions);
  const [selectedRows, setSelectedRows] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
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

    setLoading(true);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const dataUri = e.target?.result as string;
        if (!dataUri) {
          throw new Error("No se pudo leer el archivo.");
        }
        
        const result = await extractRetenciones(dataUri);
        
        toast({
          title: "Proceso completado",
          description: `La IA extrajo los datos y la descarga del archivo .txt comenzará.`,
        });

        handleDownload(result, file.name);

      } catch (error: any) {
        console.error("Error processing PDF file:", error);
        toast({
          title: "Error al procesar el archivo con IA",
          description: error.message || "No se pudo extraer la información del PDF. Intenta con otro archivo.",
          variant: "destructive",
        });
      } finally {
        setLoading(false);
        if(fileInputRef.current) fileInputRef.current.value = "";
      }
    };
    reader.onerror = () => {
      console.error("Error reading file");
      toast({
          title: "Error al leer el archivo",
          description: "Ocurrió un error al intentar leer el archivo.",
          variant: "destructive",
      });
      setLoading(false);
    }
    reader.readAsDataURL(file);
  };
  
  const handleDownload = (content: string, originalFileName: string) => {
    if (!content) return;
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const txtFileName = originalFileName.replace(/\.[^/.]+$/, "") + ".txt";
    link.download = txtFileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedRows(retentions.map(r => r.id));
    } else {
      setSelectedRows([]);
    }
  };

  const handleSelectRow = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedRows(prev => [...prev, id]);
    } else {
      setSelectedRows(prev => prev.filter(rowId => rowId !== id));
    }
  };

  const isAllSelected = selectedRows.length === retentions.length && retentions.length > 0;

  return (
    <main className="container mx-auto px-4 py-12 space-y-8">
      <div>
        <Link href="/" className="inline-flex items-center text-sm font-medium text-primary hover:underline mb-4">
          <ArrowLeft className="mr-2 h-4 w-4" />
          Volver al portal
        </Link>
        <h1 className="text-4xl font-headline font-bold tracking-tight text-foreground">
          Historial de Retenciones (Beta)
        </h1>
        <p className="mt-2 text-lg text-foreground/80">
          Aquí puedes ver y gestionar todas las retenciones que has procesado. Esta herramienta está en desarrollo.
        </p>
      </div>

      <Card>
        <CardContent className="p-6">
          <div className="flex items-center gap-2 mb-4">
            <Button variant="outline" disabled={selectedRows.length === 0}>
              <Mail className="mr-2 h-4 w-4"/>
              Email para Anular ({selectedRows.length})
            </Button>
            <Button variant="outline" disabled={selectedRows.length === 0}>
              <Send className="mr-2 h-4 w-4"/>
              Solicitar Aceptación SRI ({selectedRows.length})
            </Button>
          </div>
          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead padding="checkbox">
                    <Checkbox
                      checked={isAllSelected}
                      onCheckedChange={(checked) => handleSelectAll(checked as boolean)}
                      aria-label="Seleccionar todo"
                    />
                  </TableHead>
                  <TableHead>Acciones Email/Copiar</TableHead>
                  <TableHead>Nro. Retención</TableHead>
                  <TableHead>Razón Social Proveedor</TableHead>
                  <TableHead>Nro. Factura</TableHead>
                  <TableHead>Valor Reten.</TableHead>
                  <TableHead>Estado</TableHead>
                  <TableHead>Fecha Creación</TableHead>
                  <TableHead>Fecha Emisión</TableHead>
                  <TableHead>Verificar SRI</TableHead>
                  <TableHead>Otras Acciones</TableHead>
                  <TableHead>Autorización</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {retentions.map((retention) => (
                  <TableRow key={retention.id}>
                    <TableCell>
                      <Checkbox
                        checked={selectedRows.includes(retention.id)}
                        onCheckedChange={(checked) => handleSelectRow(retention.id, checked as boolean)}
                        aria-label={`Seleccionar fila ${retention.id}`}
                      />
                    </TableCell>
                    <TableCell>
                      <div className="flex items-center gap-1">
                        <Button variant="ghost" size="icon" className="h-8 w-8"><Mail className="h-4 w-4"/></Button>
                        <Button variant="ghost" size="icon" className="h-8 w-8"><Send className="h-4 w-4"/></Button>
                        <Button variant="ghost" size="icon" className="h-8 w-8"><Copy className="h-4 w-4"/></Button>
                      </div>
                    </TableCell>
                    <TableCell>{retention.nroRetencion}</TableCell>
                    <TableCell className="font-medium">{retention.razonSocial}</TableCell>
                    <TableCell>{retention.nroFactura}</TableCell>
                    <TableCell>{retention.valor}</TableCell>
                    <TableCell>
                      <Badge variant={retention.estado === 'Solicitado' ? 'success' : 'destructive'}>
                        {retention.estado}
                      </Badge>
                    </TableCell>
                    <TableCell>{retention.fechaCreacion}</TableCell>
                    <TableCell>{retention.fechaEmision}</TableCell>
                    <TableCell>
                      <Button variant="outline" size="sm" asChild>
                        <a href="#" target="_blank" rel="noopener noreferrer">
                          <ExternalLink className="mr-2 h-4 w-4"/> Verificar en SRI
                        </a>
                      </Button>
                    </TableCell>
                    <TableCell>
                      <Button variant="ghost" size="icon" className="h-8 w-8 text-destructive hover:text-destructive">
                        <Trash2 className="h-4 w-4"/>
                      </Button>
                    </TableCell>
                    <TableCell>{retention.autorizacion}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        </CardContent>
      </Card>
      
      <Accordion type="multiple" className="w-full">
        <AccordionItem value="no-recibidas">
          <AccordionTrigger>Mostrar Retenciones No Recibidas (1)</AccordionTrigger>
          <AccordionContent>
            <p className="p-4 text-muted-foreground">Aquí se mostraría la lista de retenciones no recibidas.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="anuladas">
          <AccordionTrigger>Mostrar Retenciones Anuladas (31)</AccordionTrigger>
          <AccordionContent>
            <p className="p-4 text-muted-foreground">Aquí se mostraría la lista de retenciones anuladas.</p>
          </AccordionContent>
        </AccordionItem>
      </Accordion>

      <Card className="max-w-xl mx-auto shadow-lg border-2 border-dashed border-primary/50 hover:border-primary transition-colors">
        <CardHeader>
          <CardTitle>Procesar Nuevo PDF</CardTitle>
          <CardDescription>Sube un archivo de retenciones en PDF para que la IA extraiga los datos y genere un archivo .txt compatible con DIMM.</CardDescription>
        </CardHeader>
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
              <p className="text-lg font-semibold text-foreground">Haz clic para subir o arrastra y suelta</p>
              <p className="text-muted-foreground">SOLO ARCHIVOS PDF</p>
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
    </main>
  );
}
