/**
 * Carga datos del mapa de puntos de venta desde Google Sheets (export CSV).
 * Misma lógica que el proyecto Map/googleSheets.js, para uso en API routes.
 */

import { getAllDataCombined } from '@/app/lib/sheets'

const HOJAS = [
  { nombre: 'Yamila', gid: '1090663139' },
  { nombre: 'Romina', gid: '822617376' },
  { nombre: 'Gisela', gid: '1787134852' },
  { nombre: 'Fabiana', gid: '1588954480' },
  { nombre: 'Anabella', gid: '2145568967' },
]

const BASE_CSV_URL =
  process.env.GOOGLE_SHEET_CSV_URL_MAPA ||
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRRLfaWpmwj_Hl2kHFkbAjgJiypqi4CNidmKqRyUqmdRNpVKDZNIeWU9-Vg0VCUHA0YhPtNXJFIrKOr/pub?output=csv&gid='

const COLUMN_MAP: Record<string, string> = {
  'ID': 'id',
  'Estado Kiosco': 'estado',
  'Paquete': 'paquete',
  'Domicilio': 'direccion',
  'Entre Calle 1': 'entre_calle_1',
  'Entre Calle 2': 'entre_calle_2',
  'Pais': 'pais',
  'Provincia': 'provincia',
  'Partido': 'partido',
  'Localidad / Barrio': 'localidad',
  'N° Vendedor': 'num_vendedor',
  'Distribuidora': 'distribuidora',
  'Dias de atención': 'dias_atencion',
  'Horario': 'horario',
  'Escaparate': 'escaparate',
  'Ubicación': 'ubicacion',
  'Fachada puesto': 'fachada',
  'Venta productos no editoriales': 'venta_no_editorial',
  'Reparto': 'reparto',
  'Suscripciones': 'suscripciones',
  'Nombre y Apellido': 'contacto_nombre',
  'Mayor venta': 'mayor_venta',
  'Utiliza Parada Online': 'usa_parada_online',
  'Teléfono': 'telefono',
  'Correo electrónico': 'email',
  'Relevado por': 'relevado_por',
  'Observaciones': 'observaciones',
  'Comentarios': 'comentarios',
  'IMG': 'imagen',
  'Latitud': 'latitud',
  'Longitud': 'longitud',
  'LATITUD': 'latitud',
  'LONGITUD': 'longitud',
  'Lat': 'latitud',
  'Lng': 'longitud',
  'latitude': 'latitud',
  'longitude': 'longitud',
  'DISPOSITIVO': 'dispositivo',
}

const RELEVADOR_NOMBRE_A_EMAIL: Record<string, string> = {
  'Yamila': 'yamisol213@gmail.com',
  'Romina': 'rominadifrancesco1977@gmail.com',
  'Gisela': 'giselayanina1309@gmail.com',
  'Fabiana': 'chuchisbuono@gmail.com',
  'Anabella': 'anabellamarchesi@gmail.com',
}

export interface LugarMapa {
  id: string
  nombre: string
  paquete: string
  tipo: string
  descripcion: string
  latitud: number
  longitud: number
  direccion: string
  localidad: string
  partido: string
  provincia: string
  telefono: string
  email: string
  contacto_nombre: string
  dias_atencion: string
  horario: string
  estado: string
  distribuidora: string
  num_vendedor: string
  imagen: string
  escaparate: string
  ubicacion: string
  fachada: string
  venta_no_editorial: string
  reparto: string
  suscripciones: string
  mayor_venta: string
  usa_parada_online: string
  relevado_por: string
  estaAbierto: boolean | null
}

function parseCSVLine(line: string): string[] {
  const result: string[] = []
  let current = ''
  let inQuotes = false
  for (let i = 0; i < line.length; i++) {
    const char = line[i]
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"'
        i++
      } else {
        inQuotes = !inQuotes
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim())
      current = ''
    } else {
      current += char
    }
  }
  result.push(current.trim())
  return result
}

function parseCSV(csvText: string): Record<string, string>[] {
  const lines = csvText.split('\n')
  const firstLine = (lines[0] || '').replace(/^\uFEFF/, '').trim()
  const rawHeaders = parseCSVLine(firstLine)
  const headers = rawHeaders.map((h) => (typeof h === 'string' ? h.trim() : String(h)))
  const indexRelevadoPor = headers.findIndex((h) => h && String(h).toLowerCase().includes('relevado'))
  const idxRelevado = indexRelevadoPor >= 0 ? indexRelevadoPor : 25
  const data: Record<string, string>[] = []

  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue
    const values = parseCSVLine(lines[i])
    const row: Record<string, string> = {}
    headers.forEach((header, index) => {
      const mappedKey = COLUMN_MAP[header] || header
      const val = values[index]
      row[mappedKey] = (val != null ? String(val).trim() : '') || ''
    })
    if (values.length > idxRelevado && values[idxRelevado] != null) {
      const v = String(values[idxRelevado]).trim()
      if (v) row.relevado_por = v
    }
    data.push(row)
  }
  return data
}

function determinarSiAbierto(row: Record<string, string>): boolean | null {
  const estado = (row.estado || '').toLowerCase()
  if (estado.includes('cerrado')) return false
  if (estado.includes('zona peligrosa')) return false
  if (estado.includes('no se encuentra')) return false
  // Abierto, Ahora es Cafeteria y Abierto pero otro rubro se consideran "activos" para el mapa
  if (estado === 'abierto') return true
  if (estado.includes('ahora es cafeteria')) return true
  if (estado.includes('abierto pero otro rubro')) return true
  return null
}

/** Parsea coordenada aceptando coma o punto como decimal (ej. -34,6037 o -34.6037) */
function parseCoord(value: string | undefined | null): number {
  const s = (value != null ? String(value).trim() : '').replace(',', '.')
  if (!s) return NaN
  const n = parseFloat(s)
  return typeof n === 'number' && !isNaN(n) ? n : NaN
}

/** Obtiene lat/long de una fila probando varias claves posibles */
function getLatLng(row: Record<string, string>): { lat: number; lng: number } {
  const latStr = (row.latitud ?? row.Latitud ?? row.LATITUD ?? row.lat ?? '').trim()
  const lngStr = (row.longitud ?? row.Longitud ?? row.LONGITUD ?? row.lng ?? '').trim()
  return { lat: parseCoord(latStr), lng: parseCoord(lngStr) }
}

async function fetchLugaresFromHoja(hoja: { nombre: string; gid: string }) {
  try {
    const url = BASE_CSV_URL + hoja.gid
    const response = await fetch(url)
    if (!response.ok) {
      return { lugares: [] as LugarMapa[], raw: 0, sinCoordenadas: 0 }
    }
    const csvText = await response.text()
    const rawData = parseCSV(csvText)
    const raw = rawData.length
    const conCoordenadas = rawData.filter((row) => {
      const { lat, lng } = getLatLng(row)
      return !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0
    })
    const sinCoordenadas = raw - conCoordenadas.length
    const lugares: LugarMapa[] = conCoordenadas.map((row) => {
      const { lat, lng } = getLatLng(row)
      return {
      id: row.id || Math.random().toString(36).substring(2, 11),
      nombre: row.paquete || 'Sin nombre',
      paquete: row.paquete || '',
      tipo: 'kiosco',
      descripcion: [row.ubicacion, row.escaparate, row.fachada].filter(Boolean).join(' - '),
      latitud: lat,
      longitud: lng,
      direccion: row.direccion || '',
      localidad: row.localidad || '',
      partido: row.partido || '',
      provincia: row.provincia || '',
      telefono: row.telefono || '',
      email: row.email || '',
      contacto_nombre: row.contacto_nombre || '',
      dias_atencion: row.dias_atencion || '',
      horario: row.horario || '',
      estado: row.estado || 'Abierto',
      distribuidora: row.distribuidora || '',
      num_vendedor: row.num_vendedor || '',
      imagen: row.imagen || '',
      escaparate: row.escaparate || '',
      ubicacion: row.ubicacion || '',
      fachada: row.fachada || '',
      venta_no_editorial: row.venta_no_editorial || '',
      reparto: row.reparto || '',
      suscripciones: row.suscripciones || '',
      mayor_venta: row.mayor_venta || '',
      usa_parada_online: row.usa_parada_online || '',
      relevado_por:
        (row.relevado_por != null && String(row.relevado_por).trim() !== '')
          ? String(row.relevado_por).trim()
          : (RELEVADOR_NOMBRE_A_EMAIL[hoja.nombre] || hoja.nombre),
      estaAbierto: determinarSiAbierto(row),
    }
    })
    return { lugares, raw, sinCoordenadas }
  } catch {
    return { lugares: [] as LugarMapa[], raw: 0, sinCoordenadas: 0 }
  }
}

let cacheData: {
  list: LugarMapa[]
  meta: {
    totalRelevados: number
    totalConCoordenadas: number
    enMapa: number
    sinCoordenadas: number
    duplicadosEliminados: number
  }
} | null = null
let cacheTimestamp = 0
const CACHE_DURATION = 5 * 60 * 1000 // 5 min

export async function getLugaresMapa(forceRefresh = false): Promise<LugarMapa[]> {
  const now = Date.now()
  if (!forceRefresh && cacheData && now - cacheTimestamp < CACHE_DURATION) {
    return cacheData.list
  }
  const resultados = await Promise.all(HOJAS.map((hoja) => fetchLugaresFromHoja(hoja)))
  let totalRelevados = 0
  let totalSinCoordenadas = 0
  const todosLosLugares: LugarMapa[] = []
  for (const r of resultados) {
    totalRelevados += r.raw
    totalSinCoordenadas += r.sinCoordenadas
    todosLosLugares.push(...r.lugares)
  }
  const totalConCoordenadas = todosLosLugares.length
  // Contar duplicados (mismo id + mismas coords en varias hojas) solo para estadísticas.
  const idsVistos = new Set<string>()
  let duplicadosEliminados = 0
  for (const lugar of todosLosLugares) {
    const key = `${lugar.id}-${lugar.latitud}-${lugar.longitud}`
    if (idsVistos.has(key)) duplicadosEliminados++
    else idsVistos.add(key)
  }
  // Devolver todos los que tienen coordenadas (sin quitar duplicados) para que el mapa muestre
  // el mismo número que las estadísticas "relevados con coordenadas" (380 = 380).
  cacheData = {
    list: todosLosLugares,
    meta: {
      totalRelevados,
      totalConCoordenadas,
      enMapa: todosLosLugares.length,
      sinCoordenadas: totalSinCoordenadas,
      duplicadosEliminados,
    },
  }
  cacheTimestamp = now
  return cacheData.list
}

export function getLugaresMapaMeta() {
  return cacheData ? cacheData.meta : null
}

/** Índice de columna por header normalizado (minúsculas, sin espacios extra) */
function findHeaderIndex(headers: string[], ...names: string[]): number {
  const lower = headers.map((h) => String(h ?? '').trim().toLowerCase())
  for (const name of names) {
    const i = lower.findIndex((h) => h === name || h.includes(name))
    if (i !== -1) return i
  }
  return -1
}

/**
 * Obtiene los lugares con coordenadas desde el mismo spreadsheet que usa la app (API de Sheets).
 * Así el mapa muestra siempre el mismo total que "Relevados con Coordenadas" en estadísticas.
 */
export async function getLugaresDesdeMainSheet(accessToken: string): Promise<LugarMapa[]> {
  const allData = await getAllDataCombined(accessToken)
  if (!allData.length) return []
  const headers = (allData[0] || []).map((c: unknown) => String(c ?? '').trim())
  const rows = allData.slice(1) as unknown[][]

  const idxId = findHeaderIndex(headers, 'id')
  const idxRelevador = findHeaderIndex(headers, 'relevado por', 'relevador', 'censado por')
  const idxLat = findHeaderIndex(headers, 'latitud', 'lat')
  const idxLng = findHeaderIndex(headers, 'longitud', 'lng', 'long')
  const idxPaquete = findHeaderIndex(headers, 'paquete')
  const idxDomicilio = findHeaderIndex(headers, 'domicilio', 'direccion', 'dirección')
  const idxEstado = findHeaderIndex(headers, 'estado kiosco', 'estado')
  const idxPartido = findHeaderIndex(headers, 'partido')
  const idxLocalidad = findHeaderIndex(headers, 'localidad', 'localidad / barrio', 'barrio')
  const idxProvincia = findHeaderIndex(headers, 'provincia')
  const idxTelefono = findHeaderIndex(headers, 'teléfono', 'telefono')
  const idxEmail = findHeaderIndex(headers, 'correo', 'email')
  const idxContacto = findHeaderIndex(headers, 'nombre y apellido', 'contacto')
  const idxDias = findHeaderIndex(headers, 'dias de atención', 'dias de atencion')
  const idxHorario = findHeaderIndex(headers, 'horario')
  const idxDistribuidora = findHeaderIndex(headers, 'distribuidora')
  const idxNumVendedor = findHeaderIndex(headers, 'n° vendedor', 'vendedor')
  const idxImagen = findHeaderIndex(headers, 'img', 'imagen')
  const idxEscaparate = findHeaderIndex(headers, 'escaparate')
  const idxUbicacion = findHeaderIndex(headers, 'ubicación', 'ubicacion')
  const idxFachada = findHeaderIndex(headers, 'fachada')
  const idxVentaNoEdit = findHeaderIndex(headers, 'venta productos no editoriales', 'venta no editorial')
  const idxReparto = findHeaderIndex(headers, 'reparto')
  const idxSuscripciones = findHeaderIndex(headers, 'suscripciones')
  const idxMayorVenta = findHeaderIndex(headers, 'mayor venta')
  const idxParadaOnline = findHeaderIndex(headers, 'parada online', 'utiliza parada')

  if (idxRelevador === -1 || idxLat === -1 || idxLng === -1) return []

  const lugares: LugarMapa[] = []
  for (const row of rows) {
    const relevador = String(row[idxRelevador] ?? '').trim()
    if (!relevador) continue
    const lat = parseCoord(row[idxLat] as string | null | undefined)
    const lng = parseCoord(row[idxLng] as string | null | undefined)
    if (isNaN(lat) || isNaN(lng) || lat === 0 || lng === 0) continue

    const get = (i: number) => (i !== -1 && row[i] != null ? String(row[i]).trim() : '')
    const estado = get(idxEstado) || 'Abierto'
    const rowId = idxId !== -1 ? get(idxId) : String(row[0] ?? '').trim()
    const rowObj: Record<string, string> = {
      id: rowId || Math.random().toString(36).substring(2, 11),
      paquete: get(idxPaquete),
      direccion: get(idxDomicilio),
      estado,
      partido: get(idxPartido),
      localidad: get(idxLocalidad),
      provincia: get(idxProvincia),
      telefono: get(idxTelefono),
      email: get(idxEmail),
      contacto_nombre: get(idxContacto),
      dias_atencion: get(idxDias),
      horario: get(idxHorario),
      distribuidora: get(idxDistribuidora),
      num_vendedor: get(idxNumVendedor),
      imagen: get(idxImagen),
      escaparate: get(idxEscaparate),
      ubicacion: get(idxUbicacion),
      fachada: get(idxFachada),
      venta_no_editorial: get(idxVentaNoEdit),
      reparto: get(idxReparto),
      suscripciones: get(idxSuscripciones),
      mayor_venta: get(idxMayorVenta),
      usa_parada_online: get(idxParadaOnline),
      relevado_por: relevador,
    }

    lugares.push({
      id: rowObj.id,
      nombre: rowObj.paquete || 'Sin nombre',
      paquete: rowObj.paquete || '',
      tipo: 'kiosco',
      descripcion: [rowObj.ubicacion, rowObj.escaparate, rowObj.fachada].filter(Boolean).join(' - '),
      latitud: lat,
      longitud: lng,
      direccion: rowObj.direccion || '',
      localidad: rowObj.localidad || '',
      partido: rowObj.partido || '',
      provincia: rowObj.provincia || '',
      telefono: rowObj.telefono || '',
      email: rowObj.email || '',
      contacto_nombre: rowObj.contacto_nombre || '',
      dias_atencion: rowObj.dias_atencion || '',
      horario: rowObj.horario || '',
      estado: rowObj.estado || 'Abierto',
      distribuidora: rowObj.distribuidora || '',
      num_vendedor: rowObj.num_vendedor || '',
      imagen: rowObj.imagen || '',
      escaparate: rowObj.escaparate || '',
      ubicacion: rowObj.ubicacion || '',
      fachada: rowObj.fachada || '',
      venta_no_editorial: rowObj.venta_no_editorial || '',
      reparto: rowObj.reparto || '',
      suscripciones: rowObj.suscripciones || '',
      mayor_venta: rowObj.mayor_venta || '',
      usa_parada_online: rowObj.usa_parada_online || '',
      relevado_por: rowObj.relevado_por || '',
      estaAbierto: determinarSiAbierto(rowObj),
    })
  }
  return lugares
}

export function clearLugaresMapaCache() {
  cacheData = null
  cacheTimestamp = 0
}
