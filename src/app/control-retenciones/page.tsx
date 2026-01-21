"use client";

import Link from 'next/link';
import { ArrowLeft, Wrench } from 'lucide-react';
import { Button } from '@/components/ui/button';

export default function ControlRetencionesPage() {
  return (
    <main className="container mx-auto px-4 py-12">
      <div className="mb-8">
        <Link href="/" className="inline-flex items-center text-sm font-medium text-primary hover:underline">
          <ArrowLeft className="mr-2 h-4 w-4" />
          Volver al portal
        </Link>
      </div>
      <div className="flex flex-col items-center justify-center text-center py-20">
        <Wrench className="w-24 h-24 text-primary/50 mb-4" />
        <h1 className="text-4xl font-headline font-bold tracking-tight text-foreground sm:text-5xl lg:text-6xl">
          Control de retenciones
        </h1>
        <p className="mt-4 max-w-2xl mx-auto text-lg text-foreground/80">
          Este módulo está en construcción. Puedes acceder a la versión actual desde el portal.
        </p>
        <Button asChild className="mt-8">
          <Link href="/">
            Volver al inicio
          </Link>
        </Button>
      </div>
    </main>
  );
}
