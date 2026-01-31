import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getUserPermissions, saveUserPermissions, getAllPermissions } from '@/app/lib/sheets'

/**
 * GET /api/permissions
 * Obtiene los permisos del usuario actual o todos los permisos si es admin
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

    // Si es admin, retornar todos los permisos
    if (userInfo.isAdmin) {
      const allPermissions = await getAllPermissions(accessToken)
      return NextResponse.json({
        isAdmin: true,
        permissions: allPermissions,
      })
    }

    // Si no es admin, solo retornar sus propios permisos
    const allowedIds = await getUserPermissions(accessToken, userInfo.email)
    return NextResponse.json({
      isAdmin: false,
      email: userInfo.email,
      allowedIds,
    })
  } catch (error: any) {
    console.error('Error en API /api/permissions GET:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}

/**
 * POST /api/permissions
 * Asigna permisos a un usuario (solo admins)
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

    // Solo admins pueden asignar permisos
    if (!userInfo.isAdmin) {
      return NextResponse.json(
        { error: 'No autorizado. Solo administradores pueden asignar permisos.' },
        { status: 403 }
      )
    }

    // Obtener datos del body
    const body = await request.json()
    const { email, allowedIds } = body

    // Validar datos
    if (!email || typeof email !== 'string') {
      return NextResponse.json(
        { error: 'Email requerido' },
        { status: 400 }
      )
    }

    if (!Array.isArray(allowedIds)) {
      return NextResponse.json(
        { error: 'allowedIds debe ser un array' },
        { status: 400 }
      )
    }

    // Validar que los IDs sean strings
    const validIds = allowedIds.map(id => String(id).trim()).filter(id => id.length > 0)

    // Obtener hoja asignada del body (opcional)
    const assignedSheet = body.assignedSheet || ''

    // Guardar permisos
    const success = await saveUserPermissions(accessToken, email, validIds, assignedSheet)

    if (!success) {
      return NextResponse.json(
        { error: 'Error al guardar permisos' },
        { status: 500 }
      )
    }

    return NextResponse.json({
      success: true,
      message: `Permisos actualizados para ${email}`,
      permissions: {
        email,
        allowedIds: validIds,
        assignedSheet,
      },
    })
  } catch (error: any) {
    console.error('Error en API /api/permissions POST:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}

