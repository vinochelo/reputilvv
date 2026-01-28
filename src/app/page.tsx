
"use client";

import Link from 'next/link';
import { Card } from '@/components/ui/card';
import { FileText, Building, Mail, ShieldCheck, ArrowRight, Wrench, type LucideIcon } from 'lucide-react';
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
    title: 'Envío correos en masa',
    description: 'Envía correos personalizados a una lista de contactos.',
    href: 'https://correos-sigma.vercel.app/',
    icon: Mail,
  },
  {
    title: 'Control de retenciones',
    description: 'Seguimiento de retenciones anuladas',
    href: 'https://extractor-kohl.vercel.app/',
    icon: ShieldCheck,
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
    title: 'Reportes de Retail (En desarrollo)',
    description: 'Versión anterior para procesar reportes. Se moverá a producción pronto.',
    href: '/reporte-retail',
    icon: Building,
    status: 'beta',
  },
];

const ToolCard = ({ tool }: { tool: Tool }) => (
    <Link 
        href={tool.href} 
        key={tool.title} 
        className="block group h-full"
        target={tool.href.startsWith('http') ? '_blank' : undefined}
        rel={tool.href.startsWith('http') ? 'noopener noreferrer' : undefined}
    >
        <Card className="h-full flex flex-col p-6 transition-all duration-300 transform hover:-translate-y-1 hover:shadow-2xl rounded-2xl border-2 border-transparent hover:border-primary/50 bg-card/50 dark:bg-card">
            <div className="flex-grow">
                <div className={cn(
                    "bg-primary/10 p-3 rounded-full self-start inline-block mb-4", 
                    { "bg-destructive/10": tool.status === 'beta' }
                )}>
                    <tool.icon className={cn(
                        "w-6 h-6 text-primary transition-transform duration-500 group-hover:rotate-[360deg]", 
                        { "text-destructive": tool.status === 'beta' }
                    )} />
                </div>
                <h2 className="font-headline text-xl font-bold">{tool.title}</h2>
                <p className="mt-2 text-sm text-foreground/70">{tool.description}</p>
            </div>
            <div className="flex items-center text-sm font-semibold text-primary mt-6">
                Ir a la herramienta
                <ArrowRight className="ml-2 h-4 w-4 transition-transform group-hover:translate-x-1" />
            </div>
        </Card>
    </Link>
);


export default function Home() {
  return (
    <main>
      <section className="bg-muted/30 dark:bg-muted/10 border-b">
        <div className="container mx-auto px-4 py-12 sm:py-16">
          <div className="text-center">
            <div className="flex justify-center items-baseline gap-1 mb-8" aria-label="ETAFASHION RM">
              <span className="font-headline font-light text-foreground text-5xl sm:text-6xl md:text-7xl tracking-wider">
                ETAFASHION
              </span>
              <span className="font-headline font-bold text-destructive text-5xl sm:text-6xl md:text-7xl">
                RM
              </span>
            </div>
            <h1 className="text-3xl sm:text-4xl font-headline font-bold tracking-tight">
              <span className="bg-gradient-to-r from-foreground via-destructive to-foreground bg-[length:200%_auto] bg-clip-text text-transparent animate-gradient">
                Portal de Herramientas
              </span>
            </h1>
          </div>
        </div>
      </section>

      <section className="container mx-auto px-4 py-16 sm:py-20">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 max-w-4xl mx-auto">
          {mainTools.map((tool) => (
            <ToolCard tool={tool} key={tool.title} />
          ))}
        </div>
        
        <div className="max-w-4xl mx-auto mt-16">
          <Accordion type="single" collapsible className="w-full bg-card/50 dark:bg-card rounded-2xl border px-6">
            <AccordionItem value="secondary-tools" className="border-b-0">
              <AccordionTrigger className="text-lg font-semibold hover:no-underline py-6">
                  <div className="flex items-center gap-3">
                      <Wrench className="w-6 h-6 text-primary/80" />
                      <span>Herramientas Beta y de Respaldo</span>
                  </div>
              </AccordionTrigger>
              <AccordionContent>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8 pt-2 pb-6">
                  {secondaryTools.map((tool) => (
                      <ToolCard tool={tool} key={tool.title} />
                  ))}
                </div>
              </AccordionContent>
            </AccordionItem>
          </Accordion>
        </div>
      </section>
      <footer className="text-center py-8 text-sm text-foreground/60">
        <p>© 2026 Rocku. Todos los derechos reservados.</p>
      </footer>
    </main>
  );
}
