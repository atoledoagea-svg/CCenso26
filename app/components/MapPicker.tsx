'use client'

import { useEffect, useRef, useState } from 'react'
import L from 'leaflet'
import 'leaflet/dist/leaflet.css'

interface MapPickerProps {
  initialLat: number
  initialLng: number
  onCoordsChange: (lat: number, lng: number) => void
}

export default function MapPicker({ initialLat, initialLng, onCoordsChange }: MapPickerProps) {
  const mapRef = useRef<HTMLDivElement>(null)
  const mapInstanceRef = useRef<L.Map | null>(null)
  const markerRef = useRef<L.Marker | null>(null)
  const [isLoaded, setIsLoaded] = useState(false)

  useEffect(() => {
    if (!mapRef.current || mapInstanceRef.current) return

    // Fix para el icono de Leaflet en Next.js
    delete (L.Icon.Default.prototype as any)._getIconUrl
    L.Icon.Default.mergeOptions({
      iconRetinaUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon-2x.png',
      iconUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-icon.png',
      shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/images/marker-shadow.png',
    })

    // Crear mapa
    const map = L.map(mapRef.current).setView([initialLat, initialLng], 15)
    mapInstanceRef.current = map

    // Agregar capa de OpenStreetMap
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap contributors'
    }).addTo(map)

    // Crear marcador draggable
    const marker = L.marker([initialLat, initialLng], { draggable: true }).addTo(map)
    markerRef.current = marker

    // Evento cuando se arrastra el marcador
    marker.on('dragend', () => {
      const pos = marker.getLatLng()
      onCoordsChange(pos.lat, pos.lng)
    })

    // Evento cuando se hace clic en el mapa
    map.on('click', (e: L.LeafletMouseEvent) => {
      marker.setLatLng(e.latlng)
      onCoordsChange(e.latlng.lat, e.latlng.lng)
    })

    setIsLoaded(true)

    // Forzar que Leaflet recalcule el tamaño después de montar
    setTimeout(() => {
      map.invalidateSize()
    }, 100)

    // Cleanup
    return () => {
      map.remove()
      mapInstanceRef.current = null
      markerRef.current = null
    }
  }, []) // Solo inicializar una vez

  // Actualizar posición del marcador cuando cambian las coordenadas externas
  useEffect(() => {
    if (markerRef.current && mapInstanceRef.current && isLoaded) {
      const currentPos = markerRef.current.getLatLng()
      // Solo actualizar si las coordenadas son significativamente diferentes
      if (Math.abs(currentPos.lat - initialLat) > 0.0001 || Math.abs(currentPos.lng - initialLng) > 0.0001) {
        markerRef.current.setLatLng([initialLat, initialLng])
        mapInstanceRef.current.setView([initialLat, initialLng], mapInstanceRef.current.getZoom())
      }
    }
  }, [initialLat, initialLng, isLoaded])

  return (
    <div 
      ref={mapRef} 
      style={{ 
        width: '100%', 
        height: '100%',
        minHeight: '250px',
        borderRadius: '12px',
        zIndex: 1
      }} 
    />
  )
}

