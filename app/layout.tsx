import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Clarin Censo 2026',
  description: 'Clarin Censo 2026',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <head>
        <meta charSet="utf-8" />
      </head>
      <body>{children}</body>
    </html>
  )
}
