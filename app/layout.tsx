import type { Metadata, Viewport } from 'next';
import { Montserrat, JetBrains_Mono } from 'next/font/google';
import './globals.css';

const montserrat = Montserrat({
  subsets: ['latin'],
  weight: ['400', '500', '600', '700'],
  variable: '--font-montserrat',
  display: 'swap',
});

const jetbrainsMono = JetBrains_Mono({
  subsets: ['latin'],
  weight: ['400', '500', '600'],
  variable: '--font-jetbrains-mono',
  display: 'swap',
});

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  viewportFit: 'cover',
  interactiveWidget: 'resizes-content',
};

export const metadata: Metadata = {
  title: 'MK Eye Center - Patient Register Portal',
  description: 'Secure cloud-based patient register and clinic administration portal for MK Eye Center, supporting Reception, Doctor workspaces, and Google Sheets synchronization.',
  openGraph: {
    title: 'MK Excel',
    description: 'Secure cloud-based patient register and clinic administration portal for MK Eye Center.',
    type: 'website',
  },
  twitter: {
    card: 'summary_large_image',
    title: 'MK Excel',
    description: 'Secure cloud-based patient register and clinic administration portal for MK Eye Center.',
  },
};

export default function RootLayout({children}: {children: React.ReactNode}) {
  return (
    <html
      lang="en"
      className={`${montserrat.variable} ${jetbrainsMono.variable} h-full bg-slate-50 dark:bg-slate-950 antialiased`}
    >
      <body
        suppressHydrationWarning
        className={`${montserrat.className} min-h-full flex flex-col font-sans text-slate-800 dark:text-slate-100`}
      >

        {children}
      </body>
    </html>
  );
}
