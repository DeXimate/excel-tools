import type { Metadata } from 'next';
import './globals.css';
import './signature.css';
import './join-rework.css';
import './venn-clean.css';

export const metadata: Metadata = {
  title: 'ExcelFlow — Nettoyez et fusionnez vos fichiers Excel',
  description: 'Supprimez les doublons et fusionnez plusieurs fichiers Excel, directement dans votre navigateur.',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="fr">
      <body>{children}</body>
    </html>
  );
}
