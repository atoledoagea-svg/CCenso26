/**
 * Carga datos del mapa de puntos de venta desde Google Sheets (export CSV).
 * Misma lógica que el proyecto Map/googleSheets.js, para uso en API routes.
 */

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
  if (estado === 'abierto') return true
  return null
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
      const lat = parseFloat(row.latitud)
      const lng = parseFloat(row.longitud)
      return !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0
    })
    const sinCoordenadas = raw - conCoordenadas.length
    const lugares: LugarMapa[] = conCoordenadas.map((row) => ({
      id: row.id || Math.random().toString(36).substring(2, 11),
      nombre: row.paquete || 'Sin nombre',
      paquete: row.paquete || '',
      tipo: 'kiosco',
      descripcion: [row.ubicacion, row.escaparate, row.fachada].filter(Boolean).join(' - '),
      latitud: parseFloat(row.latitud),
      longitud: parseFloat(row.longitud),
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
    }))
    return { lugares, raw, sinCoordenadas }
  } catch {
    return { lugares: [] as LugarMapa[], raw: 0, sinCoordenadas: 0 }
  }
}

let cacheData: { list: LugarMapa[]; meta: { totalRelevados: number; enMapa: number; sinCoordenadas: number } } | null = null
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
  const idsVistos = new Set<string>()
  const lugaresUnicos: LugarMapa[] = []
  for (const lugar of todosLosLugares) {
    const key = `${lugar.latitud}-${lugar.longitud}`
    if (!idsVistos.has(key)) {
      idsVistos.add(key)
      lugaresUnicos.push(lugar)
    }
  }
  cacheData = {
    list: lugaresUnicos,
    meta: {
      totalRelevados,
      enMapa: lugaresUnicos.length,
      sinCoordenadas: totalSinCoordenadas,
    },
  }
  cacheTimestamp = now
  return cacheData.list
}

export function getLugaresMapaMeta() {
  return cacheData ? cacheData.meta : null
}

export function clearLugaresMapaCache() {
  cacheData = null
  cacheTimestamp = 0
}
