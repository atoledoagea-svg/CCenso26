import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getSheetsClient } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'
const ALTA_PDV_SHEET = 'ALTA PDV'
const STARTING_ID = 4279

/**
 * GET /api/alta-pdv
 * Obtiene el siguiente ID disponible para un nuevo PDV
 */
export async function GET(request: NextRequest) {
  try {
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado. Token de acceso requerido.' },
        { status: 401 }
      )
    }

    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido o expirado.' },
        { status: 401 }
      )
    }

    const sheets = getSheetsClient(accessToken)

    // Intentar obtener los IDs existentes de la hoja ALTA PDV
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${ALTA_PDV_SHEET}'!A:A`,
      })

      const rows = response.data.values || []
      
      // Encontrar el ID más alto
      let maxId = STARTING_ID
      for (let i = 1; i < rows.length; i++) { // Saltar header
        const id = parseInt(String(rows[i]?.[0] ?? '0'), 10)
        if (!isNaN(id) && id > maxId) {
          maxId = id
        }
      }

      return NextResponse.json({
        nextId: maxId + 1
      })

    } catch (error: any) {
      // Si la hoja no existe, crearla y devolver el ID inicial + 1
      if (error.code === 400 || error.message?.includes('Unable to parse range')) {
        await createAltaPdvSheet(sheets)
        return NextResponse.json({
          nextId: STARTING_ID + 1
        })
      }
      throw error
    }

  } catch (error: any) {
    console.error('Error en GET /api/alta-pdv:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error.message },
      { status: 500 }
    )
  }
}

/**
 * POST /api/alta-pdv
 * Guarda un nuevo PDV en la hoja ALTA PDV
 */
export async function POST(request: NextRequest) {
  try {
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado. Token de acceso requerido.' },
        { status: 401 }
      )
    }

    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido o expirado.' },
        { status: 401 }
      )
    }

    let body: { pdvData?: Record<string, unknown> }
    try {
      body = await request.json()
    } catch {
      return NextResponse.json(
        { error: 'Cuerpo JSON inválido' },
        { status: 400 }
      )
    }
    const { pdvData } = body

    if (!pdvData) {
      return NextResponse.json(
        { error: 'Datos del PDV requeridos' },
        { status: 400 }
      )
    }

    const sheets = getSheetsClient(accessToken)

    // Asegurar que la hoja existe
    try {
      await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${ALTA_PDV_SHEET}'!A1`,
      })
    } catch (error: any) {
      if (error.code === 400 || error.message?.includes('Unable to parse range')) {
        await createAltaPdvSheet(sheets)
      }
    }

    // Obtener el siguiente ID
    const idsResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${ALTA_PDV_SHEET}'!A:A`,
    })

    const rows = idsResponse.data.values || []
    let maxId = STARTING_ID
    for (let i = 1; i < rows.length; i++) {
      const id = parseInt(String(rows[i]?.[0] ?? '0'), 10)
      if (!isNaN(id) && id > maxId) {
        maxId = id
      }
    }
    const newId = maxId + 1

    // Fecha actual
    const today = new Date().toLocaleDateString('es-AR', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    })

    // Preparar la fila de datos en el orden especificado
    const rowData = [
      newId,                              // ID
      pdvData.estadoKiosco || '',         // Estado Kiosco
      pdvData.paquete || '',              // Paquete
      pdvData.domicilio || '',            // Domicilio
      pdvData.provincia || '',            // Provincia
      pdvData.partido || '',              // Partido
      pdvData.localidad || '',            // Localidad/Barrio
      pdvData.nVendedor || '',            // N° Vendedor
      pdvData.distribuidora || '',        // Distribuidora
      pdvData.diasAtencion || '',         // Días de atención
      pdvData.horario || '',              // Horario
      pdvData.escaparate || '',           // Escaparate
      pdvData.ubicacion || '',            // Ubicación
      pdvData.fachada || '',              // Fachada puesto
      pdvData.ventaNoEditorial || '',     // Venta productos no editoriales
      pdvData.reparto || '',              // Reparto
      pdvData.suscripciones || '',        // Suscripciones
      pdvData.nombreApellido || '',       // Nombre y Apellido
      pdvData.mayorVenta || '',           // Mayor venta
      pdvData.paradaOnline || '',         // Utiliza Parada Online
      pdvData.telefono || '',             // Teléfono
      pdvData.correoElectronico || '',    // Correo electrónico
      pdvData.observaciones || '',        // Observaciones
      pdvData.comentarios || '',          // Comentarios
      userInfo.email,                     // Relevado por
      pdvData.imageUrl || ''              // IMG
    ]

    // Agregar la fila
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${ALTA_PDV_SHEET}'!A:Z`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [rowData],
      },
    })

    return NextResponse.json({
      success: true,
      newId: newId,
      message: `PDV #${newId} creado correctamente`
    })

  } catch (error: any) {
    console.error('Error en POST /api/alta-pdv:', error)
    return NextResponse.json(
      { error: 'Error al guardar el PDV', details: error.message },
      { status: 500 }
    )
  }
}

/**
 * Crea la hoja ALTA PDV si no existe
 */
async function createAltaPdvSheet(sheets: any) {
  try {
    // Crear la hoja
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: ALTA_PDV_SHEET,
              },
            },
          },
        ],
      },
    })

    // Agregar encabezados en el orden especificado
    const headers = [
      'ID',
      'Estado Kiosco',
      'Paquete',
      'Domicilio',
      'Provincia',
      'Partido',
      'Localidad / Barrio',
      'N° Vendedor',
      'Distribuidora',
      'Dias de atención',
      'Horario',
      'Escaparate',
      'Ubicación',
      'Fachada puesto',
      'Venta productos no editoriales',
      'Reparto',
      'Suscripciones',
      'Nombre y Apellido',
      'Mayor venta',
      'Utiliza Parada Online',
      'Teléfono',
      'Correo electrónico',
      'Observaciones',
      'Comentarios',
      'Relevado por',
      'IMG'
    ]

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${ALTA_PDV_SHEET}'!A1:Z1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [headers],
      },
    })

    console.log('Hoja ALTA PDV creada correctamente')
  } catch (error) {
    console.error('Error creando hoja ALTA PDV:', error)
  }
}

