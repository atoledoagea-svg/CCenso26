import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole, type UserRole } from '@/app/lib/auth'
import { getUserPermissions, saveUserPermissions, getAllPermissions, type UserLevel } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

/**
 * GET /api/permissions
 * Obtiene los permisos del usuario actual o todos los permisos si es admin/supervisor
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

    // Obtener permisos y nivel del usuario actual
    const currentUserPermissions = await getUserPermissions(accessToken, userInfo.email)
    const role: UserRole = getUserRole(userInfo.email, currentUserPermissions.level)
    const isAdmin = role === 'admin' || role === 'supervisor'

    // Si es admin o supervisor, retornar todos los permisos
    if (isAdmin) {
      const allPermissions = await getAllPermissions(accessToken)
      return NextResponse.json({
        isAdmin: true,
        role,
        level: currentUserPermissions.level,
        permissions: allPermissions,
      })
    }

    // Si no es admin ni supervisor, solo retornar sus propios permisos
    return NextResponse.json({
      isAdmin: false,
      role,
      level: currentUserPermissions.level,
      email: userInfo.email,
      allowedIds: currentUserPermissions,
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

    // Obtener nivel del usuario actual
    const currentUserPermissions = await getUserPermissions(accessToken, userInfo.email)
    const role: UserRole = getUserRole(userInfo.email, currentUserPermissions.level)

    // Solo admins pueden asignar permisos (supervisores NO pueden)
    if (role !== 'admin') {
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

    // Obtener hoja asignada y nivel del body (opcional)
    const assignedSheet = body.assignedSheet || ''
    
    // Validar nivel (1, 2, o 3)
    let level: UserLevel = 1
    if (body.level !== undefined) {
      const parsedLevel = parseInt(String(body.level), 10)
      if (parsedLevel === 2 || parsedLevel === 3) {
        level = parsedLevel as UserLevel
      }
    }

    // Guardar permisos con nivel
    const success = await saveUserPermissions(accessToken, email, validIds, assignedSheet, level)

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
        level,
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
