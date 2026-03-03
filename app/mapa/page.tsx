'use client'

import { useEffect, useState } from 'react'

const TOKEN_KEY = 'clarin_access_token'

export default function MapaPage() {
  const [token, setToken] = useState<string | null>(null)
  const [iframeLoaded, setIframeLoaded] = useState(false)

  useEffect(() => {
    const t = typeof sessionStorage !== 'undefined' ? sessionStorage.getItem(TOKEN_KEY) : null
    if (!t) {
      window.location.href = '/'
      return
    }
    setToken(t)
  }, [])

  if (!token) {
    return (
      <div className="mapa-page mapa-page-embed">
        <div className="mapa-loading">Redirigiendo...</div>
      </div>
    )
  }

  return (
    <div className="mapa-page mapa-page-embed">
      {!iframeLoaded && (
        <div className="mapa-loading" aria-live="polite">
          Cargando mapa...
        </div>
      )}
      <iframe
        src="/mapa-app/index.html"
        className="mapa-embed-iframe"
        title="Mapa de Puntos de Venta Clarín"
        onLoad={() => setIframeLoaded(true)}
        style={{ opacity: iframeLoaded ? 1 : 0 }}
      />
      <style jsx>{`
        .mapa-page-embed {
          display: block;
          height: 100vh;
          overflow: hidden;
          position: relative;
        }
        .mapa-embed-iframe {
          width: 100%;
          height: 100%;
          border: none;
          display: block;
          transition: opacity 0.2s ease;
        }
        .mapa-loading {
          position: absolute;
          inset: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          color: #666;
          background: #f5f5f5;
        }
      `}</style>
    </div>
  )
}
