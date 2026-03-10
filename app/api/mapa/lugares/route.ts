import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole } from '@/app/lib/auth'
import { getUserPermissions } from '@/app/lib/sheets'
import { getLugaresDesdeMainSheet } from '@/app/lib/mapa-sheets'

export const dynamic = 'force-dynamic'

/**
 * GET /api/mapa/lugares
 * Lista lugares para el mapa. Solo Admin y Supervisor.
 */
export async function GET(request: NextRequest) {
  try {
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json({ error: 'No autorizado. Token requerido.' }, { status: 401 })
    }
    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json({ error: 'Token inválido o expirado.' }, { status: 401 })
    }
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const role = getUserRole(userInfo.email, userPermissions.level)
    if (role !== 'admin' && role !== 'supervisor') {
      return NextResponse.json(
        { error: 'Solo administradores y supervisores pueden ver el mapa.' },
        { status: 403 }
      )
    }
    const { estado, localidad, partido, distribuidora } = Object.fromEntries(request.nextUrl.searchParams)
    // Usar el mismo spreadsheet que las estadísticas para que el mapa muestre el mismo total que "Relevados con Coordenadas"
    let lugares = await getLugaresDesdeMainSheet(accessToken)
    if (estado) {
      const e = estado.toLowerCase()
      if (e === 'abierto') lugares = lugares.filter((l) => l.estaAbierto === true)
      else if (e === 'cerrado') lugares = lugares.filter((l) => l.estaAbierto === false)
    }
    if (localidad) {
      lugares = lugares.filter((l) =>
        (l.localidad || '').toLowerCase().includes(localidad.toLowerCase())
      )
    }
    if (partido) {
      lugares = lugares.filter((l) =>
        (l.partido || '').toLowerCase().includes(partido.toLowerCase())
      )
    }
    if (distribuidora) {
      lugares = lugares.filter((l) =>
        (l.distribuidora || '').toLowerCase().includes(distribuidora.toLowerCase())
      )
    }
    return NextResponse.json(lugares)
  } catch (error: unknown) {
    console.error('Error API mapa/lugares:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor' },
      { status: 500 }
    )
  }
}
