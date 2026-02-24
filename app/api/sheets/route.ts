import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole } from '@/app/lib/auth'
import { getAvailableSheets, getSheetIds, getUserPermissions } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

/**
 * GET /api/sheets
 * Obtiene las hojas disponibles del spreadsheet (solo para admins/supervisores)
 */
export async function GET(request: NextRequest) {
  try {
    // Verificar autenticación
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado. Token de acceso requerido.' },
        { status: 401 }
      )
    }

    // Validar token y obtener email del usuario
    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido o expirado, cierre sesión y vuelva a ingresar.' },
        { status: 401 }
      )
    }

    // Obtener nivel del usuario desde la hoja de permisos
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const role = getUserRole(userInfo.email, userPermissions.level)

    // Solo admins y supervisores pueden ver las hojas
    if (role !== 'admin' && role !== 'supervisor') {
      return NextResponse.json(
        { error: 'No autorizado. Solo administradores pueden ver las hojas.' },
        { status: 403 }
      )
    }

    // Obtener hojas disponibles
    let sheets = await getAvailableSheets(accessToken)
    
    // Supervisores NO pueden ver la hoja "test"
    if (role === 'supervisor') {
      sheets = sheets.filter(sheet => sheet.toLowerCase() !== 'test')
    }

    return NextResponse.json({
      sheets,
    })
  } catch (error: any) {
    console.error('Error en API /api/sheets GET:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}

/**
 * POST /api/sheets
 * Obtiene los IDs de una hoja específica (solo para admins/supervisores)
 */
export async function POST(request: NextRequest) {
  try {
    // Verificar autenticación
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado. Token de acceso requerido.' },
        { status: 401 }
      )
    }

    // Validar token y obtener email del usuario
    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido o expirado, cierre sesión y vuelva a ingresar.' },
        { status: 401 }
      )
    }

    // Obtener nivel del usuario desde la hoja de permisos
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const role = getUserRole(userInfo.email, userPermissions.level)

    // Solo admins y supervisores pueden obtener IDs de hojas
    if (role !== 'admin' && role !== 'supervisor') {
      return NextResponse.json(
        { error: 'No autorizado. Solo administradores pueden obtener IDs de hojas.' },
        { status: 403 }
      )
    }

    // Obtener nombre de la hoja del body
    let body: { sheetName?: string }
    try {
      body = await request.json()
    } catch {
      return NextResponse.json(
        { error: 'Cuerpo JSON inválido' },
        { status: 400 }
      )
    }
    const { sheetName } = body

    if (!sheetName || typeof sheetName !== 'string') {
      return NextResponse.json(
        { error: 'Nombre de hoja requerido' },
        { status: 400 }
      )
    }

    // Supervisores NO pueden acceder a la hoja "test"
    if (role === 'supervisor' && sheetName.toLowerCase() === 'test') {
      return NextResponse.json(
        { error: 'No autorizado para acceder a esta hoja.' },
        { status: 403 }
      )
    }

    // Obtener IDs de la hoja
    const ids = await getSheetIds(accessToken, sheetName)

    return NextResponse.json({
      sheetName,
      ids,
      count: ids.length,
    })
  } catch (error: any) {
    console.error('Error en API /api/sheets POST:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}
