import type { Metadata } from 'next'
import './globals.css'
import { ToastProvider } from './components/Toast'
import { RegisterSw } from './components/RegisterSw'

export const metadata: Metadata = {
  title: 'Clarín - Relevamiento de PDV',
  description: 'Clarín - Relevamiento de PDV',
  manifest: '/manifest.json',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="es">
      <head>
        <meta charSet="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <meta name="theme-color" content="#E31E24" />
        <meta name="mobile-web-app-capable" content="yes" />
        <meta name="apple-mobile-web-app-status-bar-style" content="default" />
        <link rel="manifest" href="/manifest.json" />
      </head>
      <body style={{ margin: 0, padding: 0 }}>
        <ToastProvider>
          {children}
          <RegisterSw />
        </ToastProvider>
      </body>
    </html>
  )
}
