'use client'

import { useEffect, useRef, useMemo } from 'react'
import L from 'leaflet'
import 'leaflet/dist/leaflet.css'

const MARKER_COLORS: Record<string, string> = {
  abierto: '#4CAF50',
  cerrado: '#E31837',
  desconocido: '#FF9800',
}

interface Lugar {
  id: string
  nombre: string
  paquete: string
  direccion: string
  localidad: string
  partido: string
  latitud: number
  longitud: number
  estaAbierto: boolean | null
}

export default function MapaView({ lugares }: { lugares: Lugar[] }) {
  const mapRef = useRef<HTMLDivElement>(null)
  const mapInstanceRef = useRef<L.Map | null>(null)
  const markersRef = useRef<L.CircleMarker[]>([])

  const bounds = useMemo(() => {
    if (lugares.length === 0) return null
    const lats = lugares.map((l) => l.latitud)
    const lngs = lugares.map((l) => l.longitud)
    return [
      [Math.min(...lats), Math.min(...lngs)],
      [Math.max(...lats), Math.max(...lngs)],
    ] as [[number, number], [number, number]]
  }, [lugares])

  useEffect(() => {
    if (!mapRef.current || lugares.length === 0) return

    if (!mapInstanceRef.current) {
      delete (L.Icon.Default.prototype as unknown as { _getIconUrl?: string })._getIconUrl
      L.Icon.Default.mergeOptions({
        iconRetinaUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon-2x.png',
        iconUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon.png',
        shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-shadow.png',
      })

      const GBA_BOUNDS: L.LatLngBoundsLiteral = [
        [-35.0, -59.0],
        [-34.3, -58.0],
      ]
      const map = L.map(mapRef.current, {
        zoomControl: false,
        maxBounds: GBA_BOUNDS,
        maxBoundsViscosity: 1.0,
        minZoom: 10,
        maxZoom: 19,
      }).setView([-34.65, -58.45], 11)

      L.control.zoom({ position: 'bottomright' }).addTo(map)
      L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
        attribution: '&copy; OpenStreetMap &copy; CARTO',
        subdomains: 'abcd',
        maxZoom: 19,
      }).addTo(map)

      mapInstanceRef.current = map
      setTimeout(() => map.invalidateSize(), 100)
    }

    const map = mapInstanceRef.current
    markersRef.current.forEach((m) => m.remove())
    markersRef.current = []

    lugares.forEach((lugar) => {
      const color =
        lugar.estaAbierto === true
          ? MARKER_COLORS.abierto
          : lugar.estaAbierto === false
            ? MARKER_COLORS.cerrado
            : MARKER_COLORS.desconocido
      const circle = L.circleMarker([lugar.latitud, lugar.longitud], {
        radius: 8,
        fillColor: color,
        color: '#fff',
        weight: 1.5,
        opacity: 1,
        fillOpacity: 0.9,
      })
      circle.bindTooltip(
        `<strong>${(lugar.paquete || lugar.nombre || 'Sin nombre').replace(/</g, '&lt;')}</strong><br/>${(lugar.direccion || '').replace(/</g, '&lt;')}<br/>${(lugar.localidad || '').replace(/</g, '&lt;')}`,
        { direction: 'top', offset: [0, -8] }
      )
      circle.addTo(map)
      markersRef.current.push(circle)
    })

    if (bounds && lugares.length > 1) {
      map.fitBounds(bounds, { padding: [20, 20], maxZoom: 14 })
    } else if (lugares.length === 1) {
      map.setView([lugares[0].latitud, lugares[0].longitud], 14)
    }
  }, [lugares, bounds])

  useEffect(() => {
    return () => {
      markersRef.current.forEach((m) => m.remove())
      markersRef.current = []
      if (mapInstanceRef.current) {
        mapInstanceRef.current.remove()
        mapInstanceRef.current = null
      }
    }
  }, [])

  return <div ref={mapRef} className="mapa-view" style={{ width: '100%', height: '100%' }} />
}
