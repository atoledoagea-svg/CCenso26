import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getAllData, getUserPermissions, getAllDataCombined } from '@/app/lib/sheets'

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
        { error: 'Token inválido o expirado, cierre sesión y vuelva a ingresar.' },
        { status: 401 }
      )
    }

    // Si es admin, obtener todos los datos (puede ser de una hoja específica)
    if (userInfo.isAdmin) {
      // Obtener parámetro de hoja de la URL (para admins)
      const { searchParams } = new URL(request.url)
      const requestedSheet = searchParams.get('sheet') || ''
      
      console.log('Cargando datos en /api/data (Admin)')
      console.log('Hoja solicitada:', requestedSheet || 'principal')
      
      // Si se pide "Todos", obtener datos combinados de todas las hojas
      let allData: any[][]
      if (requestedSheet === 'Todos') {
        console.log('Obteniendo datos combinados de todas las hojas...')
        allData = await getAllDataCombined(accessToken)
      } else {
        allData = await getAllData(accessToken, requestedSheet)
      }
      console.log('Datos obtenidos, longitud:', allData.length)

      if (allData.length === 0) {
        return NextResponse.json({
          headers: [],
          data: [],
          permissions: { allowedIds: [], isAdmin: true, assignedSheet: '', currentSheet: requestedSheet },
        })
      }

      const headers = allData[0].map((cell: any) => String(cell || ''))
      const dataRows = allData.slice(1)

      return NextResponse.json({
        headers,
        data: dataRows,
        permissions: { allowedIds: [], isAdmin: true, assignedSheet: '', currentSheet: requestedSheet },
      })
    }

    // Si no es admin, obtener permisos y hoja asignada
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const { allowedIds, assignedSheet } = userPermissions
    
    // Obtener parámetro de hoja de la URL (para usuarios comunes)
    const { searchParams } = new URL(request.url)
    const requestedSheet = searchParams.get('sheet') || ''
    
    // Usuarios comunes solo pueden acceder a su hoja asignada o ALTA PDV
    const ALTA_PDV_SHEET = 'ALTA PDV'
    let sheetToRead = assignedSheet || ''
    let currentSheet = assignedSheet || ''
    
    // Si el usuario solicita una hoja específica, validar que sea permitida
    if (requestedSheet) {
      if (requestedSheet === ALTA_PDV_SHEET) {
        sheetToRead = ALTA_PDV_SHEET
        currentSheet = ALTA_PDV_SHEET
      } else if (requestedSheet === assignedSheet) {
        sheetToRead = assignedSheet
        currentSheet = assignedSheet
      }
      // Si solicita otra hoja, usar la asignada por defecto
    }
    
    console.log(`Cargando datos en /api/data para usuario ${userInfo.email}`)
    console.log(`Hoja asignada: ${assignedSheet || 'ninguna'}`)
    console.log(`Hoja solicitada: ${requestedSheet || 'ninguna'}`)
    console.log(`Hoja a leer: ${sheetToRead || 'principal'}`)
    
    const allData = await getAllData(accessToken, sheetToRead)
    console.log('Datos obtenidos, longitud:', allData.length)

    if (allData.length === 0) {
      return NextResponse.json({
        headers: [],
        data: [],
        permissions: { allowedIds, isAdmin: false, assignedSheet, currentSheet },
      })
    }

    const headers = allData[0].map((cell: any) => String(cell || ''))
    const dataRows = allData.slice(1)

    // Si está viendo ALTA PDV o su hoja asignada, mostrar todos los datos (sin filtrar por IDs)
    if (currentSheet === ALTA_PDV_SHEET || assignedSheet) {
      return NextResponse.json({
        headers,
        data: dataRows,
        permissions: { allowedIds, isAdmin: false, assignedSheet, currentSheet },
      })
    }
    
    // Si no tiene hoja asignada pero tiene IDs, filtrar por IDs
    if (allowedIds.length === 0) {
      return NextResponse.json({
        headers,
        data: [],
        permissions: { allowedIds: [], isAdmin: false, assignedSheet: '' },
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
        assignedSheet: '',
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

