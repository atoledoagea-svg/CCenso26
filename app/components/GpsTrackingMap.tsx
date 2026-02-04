'use client'

import { useEffect, useRef } from 'react'
import L from 'leaflet'
import 'leaflet/dist/leaflet.css'

interface GpsLog {
  id: number
  fecha: string
  hora: string
  email: string
  latitud: number
  longitud: number
  dispositivo: string
  evento: string
}

interface GpsTrackingMapProps {
  logs: GpsLog[]
  selectedUser: string | null
}

export default function GpsTrackingMap({ logs, selectedUser }: GpsTrackingMapProps) {
  const mapRef = useRef<HTMLDivElement>(null)
  const mapInstanceRef = useRef<L.Map | null>(null)
  const markersRef = useRef<L.Marker[]>([])
  const polylineRef = useRef<L.Polyline | null>(null)

  useEffect(() => {
    if (!mapRef.current) return

    // Fix para el icono de Leaflet en Next.js
    delete (L.Icon.Default.prototype as any)._getIconUrl
    L.Icon.Default.mergeOptions({
      iconRetinaUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon-2x.png',
      iconUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon.png',
      shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-shadow.png',
    })

    // Crear mapa si no existe
    if (!mapInstanceRef.current) {
      const map = L.map(mapRef.current).setView([-34.6037, -58.3816], 12)
      mapInstanceRef.current = map

      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '© OpenStreetMap contributors'
      }).addTo(map)
    }

    const map = mapInstanceRef.current

    // Limpiar marcadores y líneas anteriores
    markersRef.current.forEach(marker => marker.remove())
    markersRef.current = []
    if (polylineRef.current) {
      polylineRef.current.remove()
      polylineRef.current = null
    }

    if (logs.length === 0) {
      map.setView([-34.6037, -58.3816], 12)
      return
    }

    // Crear íconos personalizados
    const createNumberedIcon = (number: number, isFirst: boolean, isLast: boolean) => {
      let color = '#3B82F6' // Azul por defecto
      if (isFirst) color = '#10B981' // Verde para el primero (más antiguo)
      if (isLast) color = '#EF4444' // Rojo para el último (más reciente)

      return L.divIcon({
        className: 'custom-marker',
        html: `<div style="
          background: ${color};
          color: white;
          width: 28px;
          height: 28px;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 12px;
          font-weight: bold;
          border: 3px solid white;
          box-shadow: 0 2px 6px rgba(0,0,0,0.3);
        ">${number}</div>`,
        iconSize: [28, 28],
        iconAnchor: [14, 14]
      })
    }

    // Ordenar logs del más antiguo al más reciente para el recorrido
    const sortedLogs = [...logs].reverse()

    // Agregar marcadores
    const points: L.LatLng[] = []
    sortedLogs.forEach((log, index) => {
      const isFirst = index === 0
      const isLast = index === sortedLogs.length - 1
      const latlng = L.latLng(log.latitud, log.longitud)
      points.push(latlng)

      const marker = L.marker(latlng, {
        icon: createNumberedIcon(index + 1, isFirst, isLast)
      })

      // Popup con información
      marker.bindPopup(`
        <div style="min-width: 200px;">
          <strong style="color: #1f2937;">${log.email}</strong>
          <hr style="margin: 8px 0; border: none; border-top: 1px solid #e5e7eb;">
          <p style="margin: 4px 0; font-size: 13px;">
            📅 <strong>Fecha:</strong> ${log.fecha}
          </p>
          <p style="margin: 4px 0; font-size: 13px;">
            🕐 <strong>Hora:</strong> ${log.hora}
          </p>
          <p style="margin: 4px 0; font-size: 13px;">
            📍 <strong>Coords:</strong> ${log.latitud.toFixed(6)}, ${log.longitud.toFixed(6)}
          </p>
          <p style="margin: 4px 0; font-size: 13px;">
            📱 <strong>Evento:</strong> ${log.evento}
          </p>
          <p style="margin: 4px 0; font-size: 13px; color: #6b7280;">
            Punto #${index + 1} de ${sortedLogs.length}
          </p>
        </div>
      `)

      marker.addTo(map)
      markersRef.current.push(marker)
    })

    // Dibujar línea de recorrido si hay más de 1 punto
    if (points.length > 1) {
      polylineRef.current = L.polyline(points, {
        color: '#6366F1',
        weight: 3,
        opacity: 0.7,
        dashArray: '10, 10'
      }).addTo(map)
    }

    // Ajustar vista para mostrar todos los puntos
    if (points.length > 0) {
      const bounds = L.latLngBounds(points)
      map.fitBounds(bounds, { padding: [50, 50] })
    }

    // Cleanup
    return () => {
      // No destruir el mapa, solo limpiar marcadores en el próximo render
    }
  }, [logs])

  // Cleanup al desmontar
  useEffect(() => {
    return () => {
      if (mapInstanceRef.current) {
        mapInstanceRef.current.remove()
        mapInstanceRef.current = null
      }
    }
  }, [])

  return (
    <div 
      ref={mapRef} 
      style={{ 
        width: '100%', 
        height: '100%',
        minHeight: '400px',
        borderRadius: '12px',
        zIndex: 1
      }} 
    />
  )
}


