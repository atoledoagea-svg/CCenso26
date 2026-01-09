import { google } from 'googleapis'

const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'
const PERMISSIONS_SHEET_NAME = 'Permisos' // Nombre de la hoja para almacenar permisos

/**
 * Obtiene el cliente autenticado de Google Sheets
 */
export function getSheetsClient(accessToken: string) {
  const oauth2Client = new google.auth.OAuth2()
  oauth2Client.setCredentials({ access_token: accessToken })

  return google.sheets({ version: 'v4', auth: oauth2Client })
}

/**
 * Obtiene el nombre de la primera hoja del spreadsheet
 */
export async function getFirstSheetName(accessToken: string): Promise<string> {
  const sheets = getSheetsClient(accessToken)
  
  const response = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title',
  })
  
  const sheetsList = response.data.sheets
  if (sheetsList && sheetsList.length > 0 && sheetsList[0].properties?.title) {
    return sheetsList[0].properties.title
  }
  
  return 'Sheet1' // Default fallback
}

/**
 * Obtiene todos los datos del sheet principal
 */
export async function getAllData(accessToken: string) {
  const sheets = getSheetsClient(accessToken)
  
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'A:AM', // Rango de columnas A a AM
  })

  return response.data.values || []
}

/**
 * Obtiene los IDs permitidos para un usuario específico desde la hoja de Permisos
 */
export async function getUserPermissions(accessToken: string, userEmail: string): Promise<string[]> {
  const sheets = getSheetsClient(accessToken)
  
  try {
    // Leer toda la hoja de Permisos
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A:B`, // Columna A: email, Columna B: IDs permitidos
    })

    const rows = response.data.values || []
    
    // Buscar el email del usuario (comparación case-insensitive)
    const userEmailLower = userEmail.toLowerCase()
    for (const row of rows) {
      if (row[0] && String(row[0]).toLowerCase() === userEmailLower) {
        // Si encuentra el usuario, parsear los IDs
        const idsString = row[1] || ''
        if (idsString.trim() === '') {
          return []
        }
        // Parsear IDs separados por comas y limpiar espacios
        return idsString.split(',').map((id: string) => id.trim()).filter((id: string) => id.length > 0)
      }
    }
    
    // Si no encuentra el usuario, retornar array vacío (sin permisos)
    return []
  } catch (error: any) {
    // Si la hoja de Permisos no existe, crearla
    if (error.code === 400 && error.message?.includes('Unable to parse range')) {
      await createPermissionsSheet(accessToken)
      return []
    }
    console.error('Error obteniendo permisos:', error)
    return []
  }
}

/**
 * Crea la hoja de Permisos si no existe
 */
async function createPermissionsSheet(accessToken: string) {
  const sheets = getSheetsClient(accessToken)
  
  try {
    // Crear la hoja
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: PERMISSIONS_SHEET_NAME,
              },
            },
          },
        ],
      },
    })

    // Agregar encabezados
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A1:B1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [['Email', 'IDs Permitidos']],
      },
    })
  } catch (error) {
    console.error('Error creando hoja de Permisos:', error)
  }
}

/**
 * Guarda o actualiza los permisos de un usuario
 */
export async function saveUserPermissions(
  accessToken: string,
  userEmail: string,
  allowedIds: string[]
): Promise<boolean> {
  const sheets = getSheetsClient(accessToken)
  
  try {
    // Primero, asegurarse de que la hoja existe
    try {
      await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PERMISSIONS_SHEET_NAME}!A1`,
      })
    } catch (error: any) {
      if (error.code === 400) {
        await createPermissionsSheet(accessToken)
      }
    }

    // Leer toda la hoja para buscar si el usuario ya existe
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A:B`,
    })

    const rows = response.data.values || []
    const userEmailLower = userEmail.toLowerCase()
    let rowIndex = -1

    // Buscar si el usuario ya existe
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]).toLowerCase() === userEmailLower) {
        rowIndex = i + 1 // +1 porque las filas empiezan en 1 (fila 1 es encabezado)
        break
      }
    }

    // Formatear IDs como string separado por comas
    const idsString = allowedIds.join(', ')

    if (rowIndex > 0) {
      // Actualizar fila existente
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PERMISSIONS_SHEET_NAME}!A${rowIndex}:B${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[userEmail, idsString]],
        },
      })
    } else {
      // Agregar nueva fila
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PERMISSIONS_SHEET_NAME}!A:B`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[userEmail, idsString]],
        },
      })
    }

    return true
  } catch (error) {
    console.error('Error guardando permisos:', error)
    return false
  }
}

/**
 * Obtiene todos los permisos (solo para admins)
 */
export async function getAllPermissions(accessToken: string): Promise<Array<{ email: string; allowedIds: string[] }>> {
  const sheets = getSheetsClient(accessToken)
  
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A:B`,
    })

    const rows = response.data.values || []
    const permissions: Array<{ email: string; allowedIds: string[] }> = []

    // Saltar el encabezado (fila 0)
    for (let i = 1; i < rows.length; i++) {
      const email = rows[i][0]
      const idsString = rows[i][1] || ''
      
      if (email) {
        const allowedIds = idsString.split(',').map((id: string) => id.trim()).filter((id: string) => id.length > 0)
        permissions.push({
          email: String(email),
          allowedIds,
        })
      }
    }

    return permissions
  } catch (error: any) {
    if (error.code === 400 && error.message?.includes('Unable to parse range')) {
      await createPermissionsSheet(accessToken)
      return []
    }
    console.error('Error obteniendo todos los permisos:', error)
    return []
  }
}

