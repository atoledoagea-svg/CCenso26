import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole, type UserRole } from '@/app/lib/auth'
import { getAllData, getUserPermissions, getAllDataCombined } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

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

    // Obtener permisos y nivel del usuario desde la hoja de permisos
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const { allowedIds, assignedSheet, level } = userPermissions
    
    // Determinar el rol basado en el nivel (super admins siempre son admin)
    const role: UserRole = getUserRole(userInfo.email, level)
    const isAdmin = role === 'admin' || role === 'supervisor'
    
    console.log(`Usuario: ${userInfo.email}, Nivel: ${level}, Rol: ${role}`)

    // Si es admin o supervisor, obtener todos los datos (puede ser de una hoja específica)
    if (isAdmin) {
      // Obtener parámetro de hoja de la URL (para admins/supervisores)
      const { searchParams } = new URL(request.url)
      let requestedSheet = searchParams.get('sheet') || ''
      
      // Supervisores NO pueden acceder a la hoja "test" - redirigir a "Todos"
      if (role === 'supervisor' && requestedSheet.toLowerCase() === 'test') {
        console.log('Supervisor intentó acceder a hoja "test" - redirigiendo a Todos')
        requestedSheet = 'Todos'
      }
      
      console.log(`Cargando datos en /api/data (${role === 'admin' ? 'Admin' : 'Supervisor'})`)
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
          permissions: { allowedIds: [], isAdmin: true, role, level, assignedSheet: '', currentSheet: requestedSheet },
        })
      }

      const firstRow = allData[0]
      const headers = (Array.isArray(firstRow) ? firstRow : []).map((cell: any) => String(cell || ''))
      const dataRows = allData.slice(1)

      return NextResponse.json({
        headers,
        data: dataRows,
        permissions: { allowedIds: [], isAdmin: true, role, level, assignedSheet: '', currentSheet: requestedSheet },
      })
    }

    // Usuario normal (nivel 1)
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
        permissions: { allowedIds, isAdmin: false, role, level, assignedSheet, currentSheet },
      })
    }

    const firstRowUser = allData[0]
    const headers = (Array.isArray(firstRowUser) ? firstRowUser : []).map((cell: any) => String(cell || ''))
    const dataRows = allData.slice(1)

    // Si está viendo ALTA PDV o su hoja asignada, mostrar todos los datos (sin filtrar por IDs)
    if (currentSheet === ALTA_PDV_SHEET || assignedSheet) {
      return NextResponse.json({
        headers,
        data: dataRows,
        permissions: { allowedIds, isAdmin: false, role, level, assignedSheet, currentSheet },
      })
    }
    
    // Si no tiene hoja asignada pero tiene IDs, filtrar por IDs
    if (allowedIds.length === 0) {
      return NextResponse.json({
        headers,
        data: [],
        permissions: { allowedIds: [], isAdmin: false, role, level, assignedSheet: '' },
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
        role,
        level,
        assignedSheet: '',
      },
    })
  } catch (error: any) {
    const message = error?.message || String(error)
    const code = error?.code || error?.response?.status
    console.error('Error en API /api/data:', message, 'code:', code, error)

    // Mensajes útiles para errores habituales de Google (Vercel/producción)
    if (code === 403 || (typeof message === 'string' && (message.includes('Permission') || message.includes('Access Not Configured') || message.includes('does not have permission')))) {
      return NextResponse.json(
        {
          error: 'No tienes acceso al documento. Verifica que el Google Sheet esté compartido con tu correo y que tu usuario figure en la hoja "Permisos".',
          details: message,
        },
        { status: 403 }
      )
    }
    if (code === 404 || (typeof message === 'string' && message.toLowerCase().includes('not found'))) {
      return NextResponse.json(
        { error: 'Documento o hoja no encontrada. Revisa que el ID del spreadsheet sea correcto.', details: message },
        { status: 404 }
      )
    }

    return NextResponse.json(
      { error: 'Error interno del servidor', details: message },
      { status: 500 }
    )
  }
}
