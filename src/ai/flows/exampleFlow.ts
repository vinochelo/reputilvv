'use server';
/**
 * @fileOverview An example flow to demonstrate AI connectivity.
 *
 * - generateGreeting - A function that generates a simple greeting.
 * - GreetingInput - The input type for the generateGreeting function.
 * - GreetingOutput - The return type for the generateGreeting function.
 */

import { ai } from '@/ai/genkit';
import { z } from 'genkit';

const GreetingInputSchema = z.object({
  name: z.string().describe('The name of the person to greet.'),
});
export type GreetingInput = z.infer<typeof GreetingInputSchema>;

const GreetingOutputSchema = z.object({
  greeting: z.string().describe('The generated greeting.'),
});
export type GreetingOutput = z.infer<typeof GreetingOutputSchema>;


export async function generateGreeting(input: GreetingInput): Promise<GreetingOutput> {
  return exampleFlow(input);
}

const prompt = ai.definePrompt({
  name: 'examplePrompt',
  input: { schema: GreetingInputSchema },
  output: { schema: GreetingOutputSchema },
  prompt: `Generate a friendly greeting for {{{name}}}.`,
});

const exampleFlow = ai.defineFlow(
  {
    name: 'exampleFlow',
    inputSchema: GreetingInputSchema,
    outputSchema: GreetingOutputSchema,
  },
  async (input) => {
    const { output } = await prompt(input);
    return output!;
  }
);
