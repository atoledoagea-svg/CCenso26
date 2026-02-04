import { NextRequest, NextResponse } from 'next/server'
import { validateGoogleToken, getAccessTokenFromRequest } from '@/app/lib/auth'
import { getSheetsClient } from '@/app/lib/sheets'

// Forzar renderizado dinámico (usa headers)
export const dynamic = 'force-dynamic'

const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'
const LOGS_SHEET_NAME = 'LOGs GPS'

/**
 * GET /api/gps-logs
 * Obtiene los logs de GPS (solo para admins)
 */
export async function GET(request: NextRequest) {
  try {
    // Verificar autenticación
    const accessToken = getAccessTokenFromRequest(request)
    if (!accessToken) {
      return NextResponse.json(
        { error: 'No autorizado' },
        { status: 401 }
      )
    }

    // Validar token y verificar que sea admin
    const userInfo = await validateGoogleToken(accessToken)
    if (!userInfo) {
      return NextResponse.json(
        { error: 'Token inválido' },
        { status: 401 }
      )
    }

    if (!userInfo.isAdmin) {
      return NextResponse.json(
        { error: 'Solo administradores pueden ver los logs de GPS' },
        { status: 403 }
      )
    }

    const sheets = getSheetsClient(accessToken)

    // Obtener filtro de usuario si se proporciona
    const { searchParams } = new URL(request.url)
    const filterEmail = searchParams.get('email')

    try {
      // Obtener todos los datos de la hoja LOGs GPS
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${LOGS_SHEET_NAME}'!A:G`,
      })

      const rows = response.data.values || []
      
      if (rows.length <= 1) {
        return NextResponse.json({
          logs: [],
          users: []
        })
      }

      // Primera fila son headers
      const headers = rows[0]
      const dataRows = rows.slice(1)

      // Convertir a objetos
      let logs = dataRows.map((row, index) => ({
        id: index,
        fecha: row[0] || '',
        hora: row[1] || '',
        email: row[2] || '',
        latitud: parseFloat(row[3]) || 0,
        longitud: parseFloat(row[4]) || 0,
        dispositivo: row[5] || '',
        evento: row[6] || ''
      })).filter(log => log.latitud !== 0 && log.longitud !== 0)

      // Obtener lista única de usuarios
      const users = [...new Set(logs.map(log => log.email))].sort()

      // Filtrar por email si se proporciona
      if (filterEmail) {
        logs = logs.filter(log => log.email.toLowerCase() === filterEmail.toLowerCase())
      }

      // Ordenar por fecha y hora (más reciente primero)
      logs.sort((a, b) => {
        const dateA = `${a.fecha} ${a.hora}`
        const dateB = `${b.fecha} ${b.hora}`
        return dateB.localeCompare(dateA)
      })

      return NextResponse.json({
        logs,
        users
      })
    } catch (error: any) {
      if (error.code === 400 || error.message?.includes('Unable to parse range')) {
        // La hoja no existe
        return NextResponse.json({
          logs: [],
          users: []
        })
      }
      throw error
    }
  } catch (error: any) {
    console.error('Error en API /api/gps-logs:', error)
    return NextResponse.json(
      { error: 'Error interno del servidor' },
      { status: 500 }
    )
  }
}

