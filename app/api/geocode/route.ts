import { NextRequest, NextResponse } from 'next/server'

/**
 * GET /api/geocode
 * Busca una dirección y devuelve coordenadas
 * Proxy para Nominatim para evitar problemas de CORS en mobile
 */
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const query = searchParams.get('q')
    
    if (!query) {
      return NextResponse.json(
        { error: 'Query parameter "q" is required' },
        { status: 400 }
      )
    }

    // Agregar CABA/Buenos Aires para limitar resultados
    const searchQuery = `${query}, CABA, Buenos Aires, Argentina`
    
    const response = await fetch(
      `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(searchQuery)}&limit=1&countrycodes=ar`,
      {
        headers: {
          'User-Agent': 'RelevamientoPDV/1.0',
          'Accept': 'application/json'
        }
      }
    )

    if (!response.ok) {
      console.error('Nominatim error:', response.status, response.statusText)
      return NextResponse.json(
        { error: 'Error en servicio de geocodificación' },
        { status: 502 }
      )
    }

    const results = await response.json()
    
    return NextResponse.json({ results })
  } catch (error: any) {
    console.error('Error en geocode API:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}

