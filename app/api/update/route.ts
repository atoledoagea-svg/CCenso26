import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest, getUserRole } from '@/app/lib/auth'
import { getSheetsClient, getUserPermissions, getFirstSheetName } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'

/**
 * POST /api/update
 * Actualiza una fila del sheet (solo si el usuario tiene permisos para ese ID)
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

    // Obtener datos del body
    const body = await request.json()
    const { rowId, values, sheetName: requestedSheetName } = body

    if (!rowId || !values || !Array.isArray(values)) {
      return NextResponse.json(
        { error: 'rowId y values requeridos' },
        { status: 400 }
      )
    }

    // Obtener permisos del usuario (incluyendo hoja asignada y nivel)
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const { allowedIds, assignedSheet, level } = userPermissions
    const role = getUserRole(userInfo.email, level)
    const isAdmin = role === 'admin' || role === 'supervisor'

    // Si no es admin/supervisor, verificar permisos
    if (!isAdmin) {
      // Obtener el ID de la fila (primera columna)
      const rowIdValue = values[0]
      if (!rowIdValue) {
        return NextResponse.json(
          { error: 'ID de fila requerido' },
          { status: 400 }
        )
      }

      // Si tiene hoja asignada, permitir editar cualquier fila de esa hoja
      // Si no tiene hoja asignada, verificar IDs permitidos
      if (!assignedSheet) {
        const rowIdString = String(rowIdValue).trim().toLowerCase()
        const allowedIdsLower = allowedIds.map(id => String(id).trim().toLowerCase())

        // Verificar que el usuario tiene permiso para editar esta fila
        if (!allowedIdsLower.includes(rowIdString)) {
          return NextResponse.json(
            { error: 'No tienes permiso para editar este registro' },
            { status: 403 }
          )
        }
      }
    }

    // CORRECCIÓN: Verificar que el ID recibido coincida con el ID de la fila
    const sheets = getSheetsClient(accessToken)

    try {
      // Determinar qué hoja usar:
      // 1. Si es admin/supervisor y especificó una hoja en la request, usar esa
      // 2. Si no es admin y tiene hoja asignada, usar la asignada
      // 3. Si es admin/supervisor y tiene "Todos" (sin hoja específica), buscar en todas las hojas
      // 4. Si no, usar la primera hoja
      let sheetName: string = ''
      let actualRowNumber = -1
      
      if (isAdmin && requestedSheetName) {
        sheetName = requestedSheetName
        console.log('Admin/Supervisor usando hoja seleccionada:', sheetName)
      } else if (!isAdmin && assignedSheet) {
        sheetName = assignedSheet
        console.log('Usuario usando hoja asignada:', sheetName)
      } else if (isAdmin) {
        // Admin/Supervisor con "Todos" seleccionado - buscar el ID en todas las hojas
        console.log('Admin/Supervisor con Todos - buscando ID en todas las hojas')
        
        // Obtener lista de hojas
        const metaResponse = await sheets.spreadsheets.get({
          spreadsheetId: SPREADSHEET_ID,
          fields: 'sheets.properties.title',
        })
        
        const sheetsList = metaResponse.data.sheets || []
        const excludedSheets = ['Permisos', 'Actividad', 'Hoja 1', 'Hoja 2']
        let availableSheets = sheetsList
          .map(sheet => sheet.properties?.title || '')
          .filter(name => name && !excludedSheets.includes(name))
        
        // Supervisores NO pueden editar en la hoja "test"
        if (role === 'supervisor') {
          availableSheets = availableSheets.filter(name => name.toLowerCase() !== 'test')
        }
        
        // Buscar el ID en cada hoja
        for (const currentSheet of availableSheets) {
          console.log(`Buscando ID "${rowId}" en hoja "${currentSheet}"...`)
          
          const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${currentSheet}'!A2:A`,
          })
          
          const ids = response.data.values || []
          for (let i = 0; i < ids.length; i++) {
            const cellId = String(ids[i]?.[0] || '').trim().toLowerCase()
            if (cellId === rowId.toLowerCase()) {
              sheetName = currentSheet
              actualRowNumber = i + 2
              console.log(`ID encontrado en hoja "${sheetName}" fila ${actualRowNumber}`)
              break
            }
          }
          
          if (actualRowNumber !== -1) break
        }
        
        if (actualRowNumber === -1) {
          console.error('ID no encontrado en ninguna hoja')
          return NextResponse.json(
            {
              error: `No se encontró el registro con ID "${rowId}" en ninguna hoja`,
              details: { searchedId: rowId, searchedSheets: availableSheets }
            },
            { status: 404 }
          )
        }
      } else {
        sheetName = await getFirstSheetName(accessToken)
        console.log('Usando hoja principal:', sheetName)
      }

      // Si ya encontramos la fila (caso admin con Todos), saltar la búsqueda
      if (actualRowNumber === -1) {
        // Buscar la fila por ID en la hoja específica
        console.log('Buscando fila con ID:', rowId, 'en hoja:', sheetName)

        const allDataResponse = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${sheetName}'!A2:AM`,
        })

        const allRows = allDataResponse.data.values || []
        console.log(`Total de filas de datos obtenidas: ${allRows.length}`)

        // Buscar la fila que tenga el ID correcto
        for (let i = 0; i < allRows.length; i++) {
          const row = allRows[i]
          if (row && row.length > 0) {
            const cellId = String(row[0] || '').trim().toLowerCase()
            if (cellId === rowId.toLowerCase()) {
              actualRowNumber = i + 2 // +2 porque empezamos desde fila 2
              console.log(`ID encontrado en fila ${actualRowNumber}`)
              break
            }
          }
        }

        if (actualRowNumber === -1) {
          console.error('ID no encontrado en el Sheet')
          return NextResponse.json(
            {
              error: `No se encontró el registro con ID "${rowId}" en la hoja "${sheetName}"`,
              details: { searchedId: rowId, sheetName: sheetName }
            },
            { status: 404 }
          )
        }
      }

      // Actualizar la fila encontrada
      const range = `'${sheetName}'!A${actualRowNumber}:AM${actualRowNumber}`
      console.log('Actualizando rango:', range)

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [values],
        },
      })

      console.log('Actualización exitosa')

    } catch (updateError: any) {
      console.error('Error en la actualización:', updateError)
      console.error('Mensaje de error:', updateError.message)
      console.error('Código de error:', updateError.code)
      throw updateError // Re-throw para que sea manejado por el catch principal
    }

    return NextResponse.json({
      success: true,
      message: 'Registro actualizado correctamente',
    })
  } catch (error: any) {
    console.error('Error en API /api/update:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}
