import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  metadataBase: new URL('https://invoice-watch.navig8-7311.chatgpt.site'),
  title: 'Invoice Watch — Invoice Report Builder',
  description: 'Filter invoice exports and download clean outstanding or completed invoice reports.',
  openGraph: {
    title: 'Invoice Watch',
    description: 'Filter. Summarize. Download.',
    images: [{ url: '/og.png', width: 1734, height: 907, alt: 'Invoice Watch report workflow' }],
  },
  twitter: {
    card: 'summary_large_image',
    title: 'Invoice Watch',
    description: 'Filter. Summarize. Download.',
    images: ['/og.png'],
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
