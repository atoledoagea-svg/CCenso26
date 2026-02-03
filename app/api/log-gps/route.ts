import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getSheetsClient } from '@/app/lib/sheets'

const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'
const LOGS_SHEET_NAME = 'LOGs GPS'

/**
 * POST /api/log-gps
 * Guarda un registro de ubicación GPS del usuario
 */
export async function POST(request: NextRequest) {
  try {
    // Verificar autenticación
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado' },
        { status: 401 }
      )
    }

    // Validar token y obtener email del usuario
    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido' },
        { status: 401 }
      )
    }

    // Obtener datos del body
    const body = await request.json()
    const { latitude, longitude, userAgent, isMobile, reason } = body

    if (!latitude || !longitude) {
      return NextResponse.json(
        { error: 'Coordenadas requeridas' },
        { status: 400 }
      )
    }

    const sheets = getSheetsClient(accessToken)

    // Verificar si la hoja existe, si no crearla
    try {
      await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${LOGS_SHEET_NAME}'!A1`,
      })
    } catch (error: any) {
      if (error.code === 400 || error.message?.includes('Unable to parse range')) {
        // La hoja no existe, crearla
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            requests: [{
              addSheet: {
                properties: {
                  title: LOGS_SHEET_NAME,
                }
              }
            }]
          }
        })

        // Agregar headers
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${LOGS_SHEET_NAME}'!A1:G1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [['Fecha', 'Hora', 'Email', 'Latitud', 'Longitud', 'Dispositivo', 'Evento']]
          }
        })
      }
    }

    // Preparar datos del log
    const now = new Date()
    const fecha = now.toLocaleDateString('es-AR', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric' 
    })
    const hora = now.toLocaleTimeString('es-AR', { 
      hour: '2-digit', 
      minute: '2-digit', 
      second: '2-digit' 
    })
    
    const dispositivo = isMobile ? 'Mobile' : 'Desktop'
    
    // Traducir el evento/reason
    const eventoMap: { [key: string]: string } = {
      'login': 'Inicio sesión',
      'foreground': 'Volvió a la app',
      'focus': 'Foco en ventana'
    }
    const evento = eventoMap[reason] || reason || 'Desconocido'
    
    const logRow = [
      fecha,
      hora,
      userInfo.email,
      latitude.toString(),
      longitude.toString(),
      dispositivo,
      evento
    ]

    // Agregar nueva fila al final
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${LOGS_SHEET_NAME}'!A:G`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [logRow]
      }
    })

    console.log(`GPS Log guardado: ${userInfo.email} - ${latitude}, ${longitude}`)

    return NextResponse.json({
      success: true,
      message: 'Log GPS guardado'
    })
  } catch (error: any) {
    console.error('Error en API /api/log-gps:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor' },
      { status: 500 }
    )
  }
}

