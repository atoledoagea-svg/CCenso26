import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole } from '@/app/lib/auth'
import { getAllSheetsData, getAvailableSheets, getUserPermissions } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

export async function GET(request: NextRequest) {
  try {
    const accessToken = getAccessTokenFromRequest(request)

    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado. Token de acceso requerido.' },
        { status: 401 }
      )
    }

    const userInfo = await validateGoogleToken(accessToken)
    
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido o expirado' },
        { status: 401 }
      )
    }

    // Obtener nivel del usuario desde la hoja de permisos
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const role = getUserRole(userInfo.email, userPermissions.level)

    // Admins y supervisores pueden ver estadísticas de todas las hojas
    if (role !== 'admin' && role !== 'supervisor') {
      return NextResponse.json(
        { error: 'No autorizado. Solo administradores y supervisores pueden ver estas estadísticas.' },
        { status: 403 }
      )
    }

    // Obtener datos de todas las hojas
    const allSheetsData = await getAllSheetsData(accessToken)
    const availableSheets = await getAvailableSheets(accessToken)

    return NextResponse.json({
      sheets: allSheetsData,
      availableSheets
    })

  } catch (error: any) {
    console.error('Error en API /api/stats:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}
