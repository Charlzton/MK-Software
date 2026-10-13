import type {Metadata} from 'next';
import './globals.css'; // Global styles

export const metadata: Metadata = {
  title: 'MK Eye Center - Patient Register Portal',
  description: 'Secure cloud-based patient register and clinic administration portal for MK Eye Center, supporting Reception, Doctor workspaces, and Google Sheets synchronization.',
  openGraph: {
    title: 'MK Eye Center - Patient Register Portal',
    description: 'Secure cloud-based patient register and clinic administration portal for MK Eye Center.',
    type: 'website',
  },
  twitter: {
    card: 'summary_large_image',
    title: 'MK Eye Center - Patient Register Portal',
    description: 'Secure cloud-based patient register and clinic administration portal for MK Eye Center.',
  },
};

export default function RootLayout({children}: {children: React.ReactNode}) {
  return (
    <html lang="en" className="h-full bg-slate-50 antialiased">
      <body suppressHydrationWarning className="min-h-full flex flex-col font-sans text-slate-800">
        {children}
      </body>
    </html>
  );
}
