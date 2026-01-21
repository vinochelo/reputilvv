"use client";

import Link from 'next/link';
import { Card, CardHeader, CardTitle, CardDescription, CardContent } from '@/components/ui/card';
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
    <main className="container mx-auto px-4 py-12">
      <div className="text-center mb-12">
        <h1 className="text-4xl font-headline font-bold tracking-tight text-foreground sm:text-5xl lg:text-6xl">
          Portal de Herramientas
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Accede a todas tus aplicaciones de trabajo desde un solo lugar.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-2 gap-8 max-w-4xl mx-auto">
        {tools.map((tool) => (
          <Link 
            href={tool.href} 
            key={tool.title} 
            className="block group"
            target={tool.href.startsWith('http') ? '_blank' : undefined}
            rel={tool.href.startsWith('http') ? 'noopener noreferrer' : undefined}
          >
            <Card className="h-full hover:border-primary transition-all duration-300 hover:shadow-lg hover:-translate-y-1">
              <CardHeader className="flex flex-row items-center gap-4">
                <div className="bg-primary/10 p-3 rounded-full">
                  <tool.icon className="w-6 h-6 text-primary" />
                </div>
                <div>
                  <CardTitle className="font-headline text-xl">{tool.title}</CardTitle>
                  <CardDescription>{tool.description}</CardDescription>
                </div>
              </CardHeader>
              <CardContent>
                <div className="flex justify-end items-center text-sm font-semibold text-primary group-hover:translate-x-1 transition-transform">
                  Ir a la herramienta <ArrowRight className="ml-2 h-4 w-4" />
                </div>
              </CardContent>
            </Card>
          </Link>
        ))}
      </div>
    </main>
  );
}
