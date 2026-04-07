
"use client";

import Link from 'next/link';
import { Card } from '@/components/ui/card';
import { FileText, Building, Mail, ShieldCheck, ArrowUpRight, Wrench, CheckCircle, type LucideIcon } from 'lucide-react';
import { cn } from '@/lib/utils';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";

type Tool = {
    title: string;
    description: string;
    href: string;
    icon: LucideIcon;
    status?: 'beta';
};

const mainTools: Tool[] = [
  {
    title: 'Reportes de venta en verde',
    description: 'Genera PDFs de reportes de utilidad Venta en Verde a partir del resumen en Excel.',
    href: '/reporte-venta-verde',
    icon: FileText,
  },
  {
    title: 'Reportes de Retail',
    description: 'Genera PDFs de reportes de utilidad Retail con datos de SAP.',
    href: 'https://reportesrespaldo.vercel.app/',
    icon: Building,
  },
  {
    title: 'Control de retenciones',
    description: 'Seguimiento de retenciones anuladas',
    href: 'https://extractor-kohl.vercel.app/',
    icon: ShieldCheck,
  },
  {
    title: 'Autorizaciones',
    description: 'Gestión y consulta de autorizaciones de documentos.',
    href: 'https://autorizaciones.vercel.app/',
    icon: CheckCircle,
  },
  {
    title: 'Envío correos en masa',
    description: 'Envía correos personalizados a una lista de contactos (Versión Sigma).',
    href: 'https://correos-sigma.vercel.app/',
    icon: Mail,
  },
];

const secondaryTools: Tool[] = [
  {
    title: 'Control de retenciones (Beta)',
    description: 'Nueva versión en desarrollo para gestionar retenciones con IA.',
    href: '/control-retenciones-beta',
    icon: ShieldCheck,
    status: 'beta',
  },
  {
    title: 'Reporte Venta en Verde (Respaldo)',
    description: 'Versión de respaldo con previsualización HTML (más lento).',
    href: '/reporte-venta-verde-beta',
    icon: FileText,
    status: 'beta',
  },
  {
    title: 'Reportes de Retail (En desarrollo)',
    description: 'Versión anterior para procesar reportes. Se moverá a producción pronto.',
    href: '/reporte-retail',
    icon: Building,
    status: 'beta',
  },
  {
    title: 'Envío correos en masa (Respaldo)',
    description: 'Versión anterior para el envío de correos.',
    href: 'https://mails-gamma.vercel.app/',
    icon: Mail,
    status: 'beta',
  },
];

const ToolCard = ({ tool }: { tool: Tool }) => (
    <Link 
        href={tool.href} 
        key={tool.title} 
        className="group block h-full outline-none"
        target={tool.href.startsWith('http') ? '_blank' : undefined}
        rel={tool.href.startsWith('http') ? 'noopener noreferrer' : undefined}
    >
        <Card className="relative h-full overflow-hidden border border-border/40 bg-background/40 backdrop-blur-xl transition-all duration-500 hover:border-primary/40 hover:shadow-[0_8px_30px_rgb(0,0,0,0.04)] hover:-translate-y-1 dark:bg-zinc-900/40 dark:hover:shadow-[0_8px_30px_rgba(59,130,246,0.1)]">
            {/* Subtle gradient overlay on hover */}
            <div className="absolute inset-0 bg-gradient-to-br from-primary/5 via-transparent to-transparent opacity-0 transition-opacity duration-500 group-hover:opacity-100" />
            
            <div className="relative p-6 flex flex-col h-full z-10">
                <div className="flex items-start justify-between mb-6">
                    <div className={cn(
                        "p-3 rounded-2xl bg-primary/10 text-primary ring-1 ring-primary/20 transition-all duration-500 group-hover:bg-primary group-hover:text-primary-foreground group-hover:ring-primary/50 group-hover:shadow-lg group-hover:shadow-primary/20",
                        { "bg-destructive/10 text-destructive ring-destructive/20 group-hover:bg-destructive group-hover:text-destructive-foreground group-hover:ring-destructive/50 group-hover:shadow-destructive/20": tool.status === 'beta' }
                    )}>
                        <tool.icon className="w-6 h-6" strokeWidth={1.5} />
                    </div>
                    {/* Arrow indicator */}
                    <div className="bg-background/80 backdrop-blur-sm p-2 rounded-full opacity-0 -translate-y-2 transition-all duration-300 group-hover:opacity-100 group-hover:translate-y-0 border border-border/50 text-foreground/50 group-hover:text-foreground shadow-sm">
                        <ArrowUpRight className="w-4 h-4" />
                    </div>
                </div>
                
                <div className="flex-grow">
                    <div className="flex items-center gap-2 mb-2">
                        <h2 className="font-semibold text-lg tracking-tight text-foreground/90 group-hover:text-foreground transition-colors">{tool.title}</h2>
                        {tool.status === 'beta' && (
                            <span className="text-[10px] font-medium uppercase tracking-wider bg-destructive/10 text-destructive px-2 py-0.5 rounded-full border border-destructive/20">
                                Beta
                            </span>
                        )}
                    </div>
                    <p className="text-sm text-muted-foreground leading-relaxed">{tool.description}</p>
                </div>
            </div>
        </Card>
    </Link>
);


export default function Home() {
  return (
    <div className="min-h-screen bg-background relative overflow-hidden selection:bg-primary/30 selection:text-primary-foreground">
      {/* Sleek background effect: Grid + Glow */}
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:24px_24px] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_0%,#000_70%,transparent_100%)] pointer-events-none" />
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[1000px] h-[400px] opacity-30 dark:opacity-20 pointer-events-none [mask-image:radial-gradient(ellipse_at_center,black,transparent_80%)]">
        <div className="absolute inset-0 bg-gradient-to-r from-primary/40 via-blue-500/40 to-purple-500/40 blur-[100px]" />
      </div>

      <div className="relative mx-auto max-w-6xl px-6 py-16 sm:py-24 lg:px-8 z-10 flex flex-col min-h-screen">
        {/* Header */}
        <header className="flex flex-col items-center text-center max-w-3xl mx-auto mb-20 space-y-6">
          <div className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full bg-muted/40 border border-border/50 text-sm font-medium text-muted-foreground backdrop-blur-sm shadow-sm">
            <span className="flex h-2 w-2 rounded-full bg-primary animate-pulse"></span>
            Workspace <span className="text-foreground/80 font-bold ml-1">ETAFASHION RM</span>
          </div>
          <h1 className="text-4xl sm:text-5xl lg:text-5xl font-bold tracking-tight text-foreground">
            Portal de <span className="bg-gradient-to-b from-primary to-primary/60 dark:from-foreground dark:to-foreground/60 bg-clip-text text-transparent">Herramientas</span>
          </h1>
          <p className="text-lg text-muted-foreground max-w-2xl leading-relaxed">
            Central de productividad y gestión. Accede rápidamente a todos tus reportes y herramientas de automatización empresariales.
          </p>
        </header>

        {/* Main Grid */}
        <main className="flex-grow w-full">
          {/* Main Tools Container (Bento style) */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 mb-12">
            {/* Make the first card span 2 cols on lg, for a slight bento effect, or just keep it symmetric. Let's keep it clean symmetric. */}
            {mainTools.map((tool) => (
              <ToolCard tool={tool} key={tool.title} />
            ))}
          </div>
          
          <div className="max-w-4xl mx-auto mt-24">
            <Accordion type="single" collapsible className="w-full">
              <AccordionItem value="secondary-tools" className="border border-border/50 rounded-2xl bg-card/30 backdrop-blur-xl overflow-hidden data-[state=open]:shadow-lg hover:border-border/80 transition-all">
                <AccordionTrigger className="text-base font-medium hover:no-underline px-6 py-5 data-[state=open]:bg-muted/20 transition-colors">
                    <div className="flex items-center gap-3">
                        <div className="p-2 rounded-xl bg-primary/10 text-primary">
                            <Wrench className="w-5 h-5" />
                        </div>
                        <span>Herramientas Secundarias y Beta</span>
                    </div>
                </AccordionTrigger>
                <AccordionContent className="px-6 pb-6 pt-4 border-t border-border/50 bg-background/20">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {secondaryTools.map((tool) => (
                        <ToolCard tool={tool} key={tool.title} />
                    ))}
                  </div>
                </AccordionContent>
              </AccordionItem>
            </Accordion>
          </div>
        </main>

        <footer className="mt-auto pt-16 pb-8 border-t border-border/40 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-muted-foreground w-full">
          <p>© {new Date().getFullYear()} Creado por Rocku. Todos los derechos reservados.</p>
          <div className="flex items-center gap-2 bg-muted/30 px-3 py-1.5 rounded-full border border-border/50">
            <span className="relative flex h-2 w-2">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
            </span>
            Sistemas Operativos
          </div>
        </footer>
      </div>
    </div>
  );
}

