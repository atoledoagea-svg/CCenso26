import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getAllData, getUserPermissions } from '@/app/lib/sheets'

/**
 * GET /api/data
 * Obtiene los datos del sheet filtrados según los permisos del usuario
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
        { error: 'Token inválido o expirado.' },
        { status: 401 }
      )
    }

    // Obtener todos los datos del sheet
    console.log('Cargando datos en /api/data')
    const allData = await getAllData(accessToken)
    console.log('Datos obtenidos, longitud:', allData.length)
    console.log('Primeras filas:', allData.slice(0, 3))

    // Si no hay datos, retornar vacío
    if (allData.length === 0) {
      return NextResponse.json({
        headers: [],
        data: [],
        permissions: { allowedIds: [], isAdmin: userInfo.isAdmin },
      })
    }

    // Separar headers (primera fila) y datos
    const headers = allData[0].map((cell: any) => String(cell || ''))
    const dataRows = allData.slice(1)

    // Si es admin, retornar todos los datos
    if (userInfo.isAdmin) {
      return NextResponse.json({
        headers,
        data: dataRows,
        permissions: { allowedIds: [], isAdmin: true },
      })
    }

    // Si no es admin, obtener IDs permitidos y filtrar
    const allowedIds = await getUserPermissions(accessToken, userInfo.email)
    
    // Si no tiene IDs asignados, retornar vacío
    if (allowedIds.length === 0) {
      return NextResponse.json({
        headers,
        data: [],
        permissions: { allowedIds: [], isAdmin: false },
      })
    }

    // Convertir IDs permitidos a minúsculas para comparación
    const allowedIdsLower = allowedIds.map(id => String(id).trim().toLowerCase())

    // Filtrar datos: solo incluir filas donde la primera columna (ID) esté en la lista permitida
    const filteredData = dataRows.filter((row: any[]) => {
      if (!row || row.length === 0) {
        return false
      }

      const firstCell = row[0]
      if (firstCell === null || firstCell === undefined) {
        return false
      }

      const rowId = String(firstCell).trim().toLowerCase()
      return allowedIdsLower.includes(rowId)
    })

    return NextResponse.json({
      headers,
      data: filteredData,
      permissions: {
        allowedIds,
        isAdmin: false,
      },
    })
  } catch (error: any) {
    console.error('Error en API /api/data:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}

