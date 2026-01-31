import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'

/**
 * POST /api/upload
 * Endpoint de respaldo - el upload principal se hace desde el cliente
 * Este endpoint solo valida que el usuario esté autenticado
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

    // Validar token
    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido o expirado.' },
        { status: 401 }
      )
    }

    // Obtener URL de la imagen (enviada desde el cliente después del upload directo)
    const body = await request.json()
    const { imageUrl } = body

    if (!imageUrl) {
      return NextResponse.json(
        { error: 'URL de imagen requerida' },
        { status: 400 }
      )
    }

    return NextResponse.json({
      success: true,
      imageUrl: imageUrl,
      verified: true,
    })

  } catch (error: any) {
    console.error('Error en API /api/upload:', error)
    return NextResponse.json(
      { error: error.message || 'Error al procesar' },
      { status: 500 }
    )
  }
}
