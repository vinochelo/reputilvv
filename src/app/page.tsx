"use client";

import Link from 'next/link';
import { Card } from '@/components/ui/card';
import { FileText, Building, Mail, ShieldCheck, ArrowRight } from 'lucide-react';

const tools = [
  {
    title: 'Reportes de venta en verde',
    description: 'Genera PDFs de utilidad a partir de archivos Excel.',
    href: '/reporte-venta-verde',
    icon: FileText,
  },
  {
    title: 'Reportes de Retail',
    description: 'Analiza y visualiza datos de ventas de retail.',
    href: '/reporte-retail',
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
    description: 'Gestiona y controla las retenciones fiscales.',
    href: 'https://extractor-kohl.vercel.app/',
    icon: ShieldCheck,
  },
];

export default function Home() {
  return (
    <main className="container mx-auto px-4 py-16 sm:py-24">
      <div className="text-center mb-16">
        <div className="flex justify-center items-start gap-1 mb-8" aria-label="ETAFASHION RM">
          <span className="font-headline font-bold text-foreground text-5xl sm:text-6xl md:text-7xl tracking-wider">
            ETAFASHION
          </span>
          <span className="font-headline font-bold text-destructive text-2xl sm:text-3xl md:text-4xl pt-1">
            RM
          </span>
        </div>
        <h1 className="text-4xl sm:text-5xl md:text-6xl font-headline font-bold tracking-tight text-foreground">
          Portal de Herramientas
        </h1>
        <p className="mt-4 max-w-3xl mx-auto text-lg text-foreground/80">
          Un espacio centralizado para acceder a todas tus aplicaciones de trabajo y mejorar tu productividad.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8 max-w-4xl mx-auto">
        {tools.map((tool) => (
          <Link 
            href={tool.href} 
            key={tool.title} 
            className="block group"
            target={tool.href.startsWith('http') ? '_blank' : undefined}
            rel={tool.href.startsWith('http') ? 'noopener noreferrer' : undefined}
          >
            <Card className="h-full flex flex-col p-8 transition-all duration-300 transform hover:-translate-y-2 hover:shadow-2xl rounded-2xl border hover:border-primary/50">
              <div className="flex-grow">
                <div className="bg-primary/10 p-4 rounded-xl self-start inline-block mb-6">
                  <tool.icon className="w-8 h-8 text-primary" />
                </div>
                <h2 className="font-headline text-2xl font-semibold">{tool.title}</h2>
                <p className="mt-2 text-base text-foreground/80">{tool.description}</p>
              </div>
              <div className="flex items-center text-sm font-bold text-primary mt-8">
                Ir a la herramienta
                <ArrowRight className="ml-2 h-5 w-5 transition-transform group-hover:translate-x-2" />
              </div>
            </Card>
          </Link>
        ))}
      </div>
    </main>
  );
}
