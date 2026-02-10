import { google } from 'googleapis'

const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'
const PERMISSIONS_SHEET_NAME = 'Permisos' // Nombre de la hoja para almacenar permisos
const DEFAULT_SHEET_NAME = 'Hoja 1' // Nombre de la hoja principal por defecto

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
 * Hojas a ocultar del panel de administración
 */
const HIDDEN_SHEETS = ['Permisos', 'Actividad', 'Hoja 2']

/**
 * Obtiene todas las hojas del spreadsheet (excluyendo las ocultas)
 */
export async function getAvailableSheets(accessToken: string): Promise<string[]> {
  const sheets = getSheetsClient(accessToken)
  
  const response = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title',
  })
  
  const sheetsList = response.data.sheets || []
  const sheetNames = sheetsList
    .map(sheet => sheet.properties?.title || '')
    .filter(name => name && !HIDDEN_SHEETS.includes(name))
  
  return sheetNames
}

/**
 * Obtiene todos los IDs de una hoja específica (primera columna)
 */
export async function getSheetIds(accessToken: string, sheetName: string): Promise<string[]> {
  const sheets = getSheetsClient(accessToken)
  
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A:A`, // Solo columna A (IDs)
    })
    
    const rows = response.data.values || []
    // Saltar el encabezado (fila 0) y obtener los IDs
    const ids = rows.slice(1)
      .map(row => String(row[0] || '').trim())
      .filter(id => id.length > 0)
    
    return ids
  } catch (error) {
    console.error(`Error obteniendo IDs de hoja ${sheetName}:`, error)
    return []
  }
}

/**
 * Obtiene todos los datos de una hoja específica o de la primera hoja disponible
 */
export async function getAllData(accessToken: string, sheetName: string = '') {
  const sheets = getSheetsClient(accessToken)
  
  // Si se especifica una hoja, usar esa
  let targetSheet = sheetName
  
  // Si no se especifica hoja, obtener la primera hoja disponible (no excluida)
  if (!targetSheet) {
    const availableSheets = await getAvailableSheets(accessToken)
    if (availableSheets.length > 0) {
      targetSheet = availableSheets[0]
    } else {
      // Fallback a la primera hoja del spreadsheet
      targetSheet = await getFirstSheetName(accessToken)
    }
  }
  
  const range = `'${targetSheet}'!A:AM`
  console.log('Cargando datos de hoja:', targetSheet)
  
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range,
  })

  return response.data.values || []
}

/**
 * Hojas a excluir del cálculo de estadísticas y de la vista "Todos"
 * Solo se incluyen las hojas de trabajo: YAMILA, ROMINA, GISELA, FABIANA, ANABELLA
 */
const STATS_EXCLUDED_SHEETS = ['Permisos', 'Actividad', 'Hoja 2', 'Hoja 1', 'LOGs GPS', 'ALTA PDV']

/**
 * Obtiene todos los datos combinados de todas las hojas (excepto las excluidas y Hoja 1)
 * Los datos se ordenan por ID (primera columna)
 */
export async function getAllDataCombined(accessToken: string): Promise<any[][]> {
  const sheets = getSheetsClient(accessToken)
  
  // Obtener lista de todas las hojas
  const metaResponse = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title',
  })
  
  const sheetsList = metaResponse.data.sheets || []
  const sheetNames = sheetsList
    .map(sheet => sheet.properties?.title || '')
    .filter(name => name && !STATS_EXCLUDED_SHEETS.includes(name))
  
  let combinedData: any[][] = []
  let headers: any[] = []
  
  for (const sheetName of sheetNames) {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!A:AM`,
      })
      const data = response.data.values || []
      
      if (data.length > 0) {
        // Guardar los headers de la primera hoja y agregar columna HOJA
        if (headers.length === 0) {
          headers = [...data[0], '__HOJA__']
        }
        // Agregar los datos (sin headers) al combinado, incluyendo el nombre de la hoja
        const rowsWithSheetName = data.slice(1).map(row => [...row, sheetName])
        combinedData = combinedData.concat(rowsWithSheetName)
      }
    } catch (error) {
      console.error(`Error obteniendo datos de hoja ${sheetName}:`, error)
    }
  }
  
  // Ordenar por ID (primera columna) de forma numérica
  combinedData.sort((a, b) => {
    const idA = parseInt(String(a[0] || '0'), 10)
    const idB = parseInt(String(b[0] || '0'), 10)
    return idA - idB
  })
  
  // Devolver con headers al inicio
  return [headers, ...combinedData]
}

/**
 * Obtiene datos de todas las hojas disponibles para estadísticas
 */
export async function getAllSheetsData(accessToken: string): Promise<{ [sheetName: string]: any[][] }> {
  const sheets = getSheetsClient(accessToken)
  
  // Obtener lista de todas las hojas
  const metaResponse = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title',
  })
  
  const sheetsList = metaResponse.data.sheets || []
  const sheetNames = sheetsList
    .map(sheet => sheet.properties?.title || '')
    .filter(name => name && !STATS_EXCLUDED_SHEETS.includes(name))
  
  // Obtener datos de cada hoja
  const allData: { [sheetName: string]: any[][] } = {}
  
  for (const sheetName of sheetNames) {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!A:AM`,
      })
      allData[sheetName] = response.data.values || []
    } catch (error) {
      console.error(`Error obteniendo datos de hoja ${sheetName}:`, error)
      allData[sheetName] = []
    }
  }
  
  return allData
}

/**
 * Niveles de usuario:
 * 1 = Usuario normal (acceso básico)
 * 2 = Supervisor (acceso intermedio, sin GPS de usuarios)
 * 3 = Admin (acceso completo)
 */
export type UserLevel = 1 | 2 | 3

/**
 * Obtiene los IDs permitidos, hoja asignada y nivel para un usuario específico desde la hoja de Permisos
 */
export async function getUserPermissions(accessToken: string, userEmail: string): Promise<{ allowedIds: string[], assignedSheet: string, level: UserLevel }> {
  const sheets = getSheetsClient(accessToken)
  
  try {
    // Leer toda la hoja de Permisos (A: email, B: IDs, C: Hoja asignada, D: Nivel)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A:D`,
    })

    const rows = response.data.values || []
    
    // Buscar el email del usuario (comparación case-insensitive)
    const userEmailLower = userEmail.toLowerCase()
    for (const row of rows) {
      if (row[0] && String(row[0]).toLowerCase() === userEmailLower) {
        // Si encuentra el usuario, parsear los IDs, la hoja asignada y el nivel
        const idsString = row[1] || ''
        const assignedSheet = row[2] || '' // Columna C: hoja asignada
        const levelValue = parseInt(String(row[3] || '1'), 10) // Columna D: nivel (default 1)
        
        // Validar que el nivel sea 1, 2 o 3
        const level: UserLevel = (levelValue === 2 || levelValue === 3) ? levelValue as UserLevel : 1
        
        const allowedIds = idsString.trim() === '' 
          ? [] 
          : idsString.split(',').map((id: string) => id.trim()).filter((id: string) => id.length > 0)
        
        return { allowedIds, assignedSheet, level }
      }
    }
    
    // Si no encuentra el usuario, retornar nivel 1 (usuario normal)
    return { allowedIds: [], assignedSheet: '', level: 1 }
  } catch (error: any) {
    // Si la hoja de Permisos no existe, crearla
    if (error.code === 400 && error.message?.includes('Unable to parse range')) {
      await createPermissionsSheet(accessToken)
      return { allowedIds: [], assignedSheet: '', level: 1 }
    }
    console.error('Error obteniendo permisos:', error)
    return { allowedIds: [], assignedSheet: '', level: 1 }
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

    // Agregar encabezados (incluyendo columna Nivel)
    // Nivel: 1=Usuario, 2=Supervisor, 3=Admin
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A1:D1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [['Email', 'IDs Permitidos', 'Hoja Asignada', 'Nivel']],
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
  allowedIds: string[],
  assignedSheet: string = '',
  level: UserLevel = 1 // Nivel del usuario: 1=Usuario, 2=Supervisor, 3=Admin
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
      range: `${PERMISSIONS_SHEET_NAME}!A:D`,
    })

    const rows = response.data.values || []
    const userEmailLower = userEmail.toLowerCase()
    let rowIndex = -1
    let existingLevel = level // Para usuarios nuevos, usar el nivel proporcionado

    // Buscar si el usuario ya existe
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]).toLowerCase() === userEmailLower) {
        rowIndex = i + 1 // +1 porque las filas empiezan en 1 (fila 1 es encabezado)
        // SIEMPRE preservar el nivel existente desde la hoja de Sheets
        // El nivel solo se puede cambiar editando directamente la hoja
        if (rows[i][3]) {
          const parsedLevel = parseInt(String(rows[i][3]), 10)
          if (parsedLevel === 1 || parsedLevel === 2 || parsedLevel === 3) {
            existingLevel = parsedLevel as UserLevel
          }
        }
        break
      }
    }

    // Formatear IDs como string separado por comas
    const idsString = allowedIds.join(', ')

    if (rowIndex > 0) {
      // Actualizar fila existente
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PERMISSIONS_SHEET_NAME}!A${rowIndex}:D${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[userEmail, idsString, assignedSheet, existingLevel]],
        },
      })
    } else {
      // Agregar nueva fila
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PERMISSIONS_SHEET_NAME}!A:D`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[userEmail, idsString, assignedSheet, level]],
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
 * Obtiene todos los permisos (solo para admins/supervisores)
 */
export async function getAllPermissions(accessToken: string): Promise<Array<{ email: string; allowedIds: string[]; assignedSheet: string; level: UserLevel }>> {
  const sheets = getSheetsClient(accessToken)
  
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${PERMISSIONS_SHEET_NAME}!A:D`,
    })

    const rows = response.data.values || []
    const permissions: Array<{ email: string; allowedIds: string[]; assignedSheet: string; level: UserLevel }> = []

    // Saltar el encabezado (fila 0)
    for (let i = 1; i < rows.length; i++) {
      const email = rows[i][0]
      const idsString = rows[i][1] || ''
      const assignedSheet = rows[i][2] || ''
      const levelValue = parseInt(String(rows[i][3] || '1'), 10)
      
      // Validar que el nivel sea 1, 2 o 3
      const level: UserLevel = (levelValue === 2 || levelValue === 3) ? levelValue as UserLevel : 1
      
      if (email) {
        const allowedIds = idsString.split(',').map((id: string) => id.trim()).filter((id: string) => id.length > 0)
        permissions.push({
          email: String(email),
          allowedIds,
          assignedSheet: String(assignedSheet),
          level,
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

