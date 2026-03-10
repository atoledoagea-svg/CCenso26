import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole } from '@/app/lib/auth'
import { getUserPermissions } from '@/app/lib/sheets'
import { getLugaresDesdeMainSheet } from '@/app/lib/mapa-sheets'

export const dynamic = 'force-dynamic'

/**
 * GET /api/mapa/stats
 * Estadísticas de relevados vs en mapa. Solo Admin y Supervisor.
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
    const lugares = await getLugaresDesdeMainSheet(accessToken)
    const total = lugares.length
    return NextResponse.json({
      totalRelevados: total,
      totalConCoordenadas: total,
      enMapa: total,
      sinCoordenadas: 0,
      duplicadosEliminados: 0,
    })
  } catch (error: unknown) {
    console.error('Error API mapa/stats:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor' },
      { status: 500 }
    )
  }
}
