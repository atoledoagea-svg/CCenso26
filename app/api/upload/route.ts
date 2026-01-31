import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'

const IMGBB_API_KEY = '4e506a27d18a2e3331ce5ffb73e20e7f'

/**
 * POST /api/upload
 * Sube una imagen a ImgBB y devuelve la URL
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

    // Obtener la imagen del form data
    const formData = await request.formData()
    const imageFile = formData.get('image') as File | null

    if (!imageFile) {
      return NextResponse.json(
        { error: 'Imagen requerida' },
        { status: 400 }
      )
    }

    // Convertir la imagen a base64
    const bytes = await imageFile.arrayBuffer()
    const buffer = Buffer.from(bytes)
    const base64Image = buffer.toString('base64')

    // Subir a ImgBB
    const imgbbFormData = new FormData()
    imgbbFormData.append('key', IMGBB_API_KEY)
    imgbbFormData.append('image', base64Image)
    imgbbFormData.append('name', `pdv_${Date.now()}`)

    const imgbbResponse = await fetch('https://api.imgbb.com/1/upload', {
      method: 'POST',
      body: imgbbFormData
    })

    if (!imgbbResponse.ok) {
      const errorText = await imgbbResponse.text()
      console.error('Error de ImgBB:', errorText)
      return NextResponse.json(
        { error: 'Error al subir la imagen a ImgBB' },
        { status: 500 }
      )
    }

    const imgbbData = await imgbbResponse.json()

    if (!imgbbData.success) {
      console.error('ImgBB no exitoso:', imgbbData)
      return NextResponse.json(
        { error: 'Error en la respuesta de ImgBB' },
        { status: 500 }
      )
    }

    // Devolver la URL de la imagen
    return NextResponse.json({
      success: true,
      imageUrl: imgbbData.data.url,
      displayUrl: imgbbData.data.display_url,
      deleteUrl: imgbbData.data.delete_url,
      thumbnail: imgbbData.data.thumb?.url || imgbbData.data.url
    })

  } catch (error: any) {
    console.error('Error en API /api/upload:', error)
    return NextResponse.json(
      { error: error.message || 'Error al procesar la imagen' },
      { status: 500 }
    )
  }
}
