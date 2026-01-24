import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getSheetsClient, getUserPermissions, getFirstSheetName } from '@/app/lib/sheets'

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
    const { rowId, values } = body

    if (!rowId || !values || !Array.isArray(values)) {
      return NextResponse.json(
        { error: 'rowId y values requeridos' },
        { status: 400 }
      )
    }

    // Obtener permisos del usuario (incluyendo hoja asignada)
    const userPermissions = await getUserPermissions(accessToken, userInfo.email)
    const { allowedIds, assignedSheet } = userPermissions

    // Si no es admin, verificar permisos
    if (!userInfo.isAdmin) {
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
      // Determinar qué hoja usar: la asignada del usuario o la primera hoja (para admin)
      let sheetName: string
      if (!userInfo.isAdmin && assignedSheet) {
        sheetName = assignedSheet
        console.log('Usando hoja asignada:', sheetName)
      } else {
        sheetName = await getFirstSheetName(accessToken)
        console.log('Usando hoja principal:', sheetName)
      }

      // Buscar la fila por ID en lugar de usar rowNumber
      console.log('Buscando fila con ID:', rowId)

      // Obtener todas las filas de datos (empezando desde la fila 2, saltando headers)
      const allDataResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A2:AM`,
      })

      const allRows = allDataResponse.data.values || []
      console.log(`Total de filas de datos obtenidas: ${allRows.length}`)

      // Buscar la fila que tenga el ID correcto
      let actualRowNumber = -1
      for (let i = 0; i < allRows.length; i++) {
        const row = allRows[i]
        if (row && row.length > 0) {
          const cellId = String(row[0] || '').trim().toLowerCase()
          console.log(`Comparando fila ${i + 2}: "${cellId}" con "${rowId.toLowerCase()}"`)
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
            error: `No se encontró el registro con ID "${rowId}" en el Sheet`,
            details: {
              searchedId: rowId,
              totalRows: allRows.length,
              sheetName: sheetName
            }
          },
          { status: 404 }
        )
      }

      // Actualizar la fila encontrada
      const range = `${sheetName}!A${actualRowNumber}:AM${actualRowNumber}`
      console.log('Actualizando rango:', range, 'con valores:', values)

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

