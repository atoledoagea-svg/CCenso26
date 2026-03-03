import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getSheetsClient } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

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
    let body: { latitude?: number; longitude?: number; userAgent?: string; isMobile?: boolean; reason?: string }
    try {
      body = await request.json()
    } catch {
      return NextResponse.json(
        { error: 'Cuerpo JSON inválido' },
        { status: 400 }
      )
    }
    const { latitude, longitude, userAgent, isMobile, reason } = body

    const lat = Number(latitude)
    const lng = Number(longitude)
    if (!Number.isFinite(lat) || lat < -90 || lat > 90 || !Number.isFinite(lng) || lng < -180 || lng > 180) {
      return NextResponse.json(
        { error: 'Coordenadas inválidas. Latitud debe estar entre -90 y 90, longitud entre -180 y 180.' },
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
        try {
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
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${LOGS_SHEET_NAME}'!A1:G1`,
            valueInputOption: 'USER_ENTERED',
            requestBody: {
              values: [['Fecha', 'Hora', 'Email', 'Latitud', 'Longitud', 'Dispositivo', 'Evento']]
            }
          })
        } catch (createError: any) {
          console.error('Error creando hoja LOGs GPS:', createError)
          return NextResponse.json(
            { error: 'Error al crear hoja de logs GPS', details: createError?.message },
            { status: 500 }
          )
        }
      } else {
        console.error('Error accediendo a hoja LOGs GPS:', error)
        return NextResponse.json(
          { error: 'Error al acceder a la hoja de logs', details: error?.message },
          { status: 500 }
        )
      }
    }

    // Preparar datos del log (zona horaria Argentina)
    const now = new Date()
    const fecha = now.toLocaleDateString('es-AR', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric',
      timeZone: 'America/Argentina/Buenos_Aires'
    })
    const hora = now.toLocaleTimeString('es-AR', { 
      hour: '2-digit', 
      minute: '2-digit', 
      second: '2-digit',
      timeZone: 'America/Argentina/Buenos_Aires'
    })
    
    const dispositivo = isMobile ? 'Mobile' : 'Desktop'
    
    // Traducir el evento/reason
    const eventoMap: { [key: string]: string } = {
      'login': 'Inicio sesión',
      'foreground': 'Volvió a la app',
      'focus': 'Foco en ventana'
    }
    const reasonKey = reason ?? ''
    const evento = eventoMap[reasonKey] || reason || 'Desconocido'
    
    const logRow = [
      fecha,
      hora,
      userInfo.email,
      lat.toString(),
      lng.toString(),
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

    return NextResponse.json({
      success: true,
      message: 'Log GPS guardado'
    })
  } catch (error: any) {
    console.error('Error en API /api/log-gps:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor', details: error?.message },
      { status: 500 }
    )
  }
}

