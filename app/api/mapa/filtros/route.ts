import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole } from '@/app/lib/auth'
import { getUserPermissions } from '@/app/lib/sheets'
import { getLugaresMapa } from '@/app/lib/mapa-sheets'

export const dynamic = 'force-dynamic'

/**
 * GET /api/mapa/filtros
 * Valores únicos para filtros del mapa. Solo Admin y Supervisor.
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
    const lugares = await getLugaresMapa()
    const localidades = [...new Set(lugares.map((l) => l.localidad).filter(Boolean))].sort()
    const partidos = [...new Set(lugares.map((l) => l.partido).filter(Boolean))].sort()
    const distribuidoras = [...new Set(lugares.map((l) => l.distribuidora).filter(Boolean))].sort()
    return NextResponse.json({ localidades, partidos, distribuidoras })
  } catch (error: unknown) {
    console.error('Error API mapa/filtros:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor' },
      { status: 500 }
    )
  }
}
