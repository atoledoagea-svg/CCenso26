'use client'

import { useEffect, useState, useCallback, useRef } from 'react'
import Image from 'next/image'
import Link from 'next/link'
import dynamic from 'next/dynamic'
import { useToast } from './components/Toast'
import { usePermissions, type Permission, type UserLevel } from './hooks/usePermissions'
import { UserCard } from './components/UserCard'
import { LoginScreen } from './components/LoginScreen'

// Importar MapPicker dinámicamente (Leaflet necesita window)
const MapPicker = dynamic(() => import('./components/MapPicker'), { 
  ssr: false,
  loading: () => <div className="map-loading">Cargando mapa...</div>
})

// Importar GpsTrackingMap dinámicamente (Leaflet necesita window)
const GpsTrackingMap = dynamic(() => import('./components/GpsTrackingMap'), { 
  ssr: false,
  loading: () => <div className="map-loading">Cargando mapa de seguimiento...</div>
})

declare global {
  interface Window {
    gapi: any
    google: any
  }
}

// Tipos de roles de usuario
type UserRole = 'user' | 'supervisor' | 'admin'

// Estado del puesto (selector en modal de edición)
type PuestoStatusValue = 'abierto' | 'cerrado' | 'no_encontrado' | 'zona_peligrosa' | 'cafeteria'

interface SheetData {
  headers: string[]
  data: any[][]
  permissions: {
    allowedIds: string[]
    isAdmin: boolean
    role?: UserRole
    level?: UserLevel
    assignedSheet?: string
    currentSheet?: string
  }
}

export default function Home() {
  const toast = useToast()
  const [tokenClient, setTokenClient] = useState<any>(null)
  const [gapiInited, setGapiInited] = useState(false)
  const [gisInited, setGisInited] = useState(false)
  
  const [isAuthorized, setIsAuthorized] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const [userEmail, setUserEmail] = useState<string | null>(null)
  const [accessToken, setAccessToken] = useState<string | null>(null)
  const [sessionExpired, setSessionExpired] = useState(false)
  const [sessionChecked, setSessionChecked] = useState(false)
  
  // Tips de uso para mostrar durante la carga
  const loadingTips = [
    "Puedes usar el botón 📍 para guardar la ubicación del PDV rápidamente.",
    "El botón de Clarín en la barra inferior recarga los datos y te lleva al inicio.",
    "Desde 'Ayuda' puedes contactar soporte por WhatsApp.",
    "Puedes buscar PDVs por ID o por nombre de paquete.",
    "Al tomar una foto con la cámara, se guardan automáticamente las coordenadas GPS.",
    "Usa 'Agregar localización' para marcar manualmente la ubicación en el mapa.",
    "Los campos con * son obligatorios y no pueden quedar vacíos.",
    "Puedes descargar el cuestionario en PDF desde el botón 'Ayuda'.",
    "El progreso de relevamiento se muestra en las estadísticas.",
    "Recuerda guardar los cambios antes de cerrar el formulario de edición."
  ]
  
  // Dashboard states
  const [sheetData, setSheetData] = useState<SheetData | null>(null)
  const [loadingData, setLoadingData] = useState(false)
  const [currentTip, setCurrentTip] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [editingRow, setEditingRow] = useState<number | null>(null)
  const [editedValues, setEditedValues] = useState<any[]>([])
  const [originalValues, setOriginalValues] = useState<any[]>([]) // Valores originales del Excel
  const [saving, setSaving] = useState(false)
  const [searchTerm, setSearchTerm] = useState('')
  const [searchType, setSearchType] = useState<'id' | 'paquete' | 'direccion'>('id')
  const [filterRelevado, setFilterRelevado] = useState<'todos' | 'relevados' | 'no_relevados' | 'censados_sin_mapear' | 'censados_mapeados'>('todos')
  const [filterPartido, setFilterPartido] = useState<string>('')
  const [filterLocalidad, setFilterLocalidad] = useState<string>('')
  
  // Ordenamiento de columnas (tipo Excel)
  const [sortColumn, setSortColumn] = useState<number | null>(null)
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc')
  
  const [showStats, setShowStats] = useState(false)
  
  // Stats by sheets states (for admin)
  const [allSheetsStats, setAllSheetsStats] = useState<{ [sheetName: string]: any[][] }>({})
  const [statsSelectedSheet, setStatsSelectedSheet] = useState<string>('Total')
  const [loadingStats, setLoadingStats] = useState(false)
  
  // Admin panel states
  const [showPermissionsPanel, setShowPermissionsPanel] = useState(false)
  const [editingPermission, setEditingPermission] = useState<Permission | null>(null)
  const [newPermIds, setNewPermIds] = useState('')
  const [savingPermissions, setSavingPermissions] = useState(false)
  const [newUserEmail, setNewUserEmail] = useState('')
  const [expandedIdsIndex, setExpandedIdsIndex] = useState<number | null>(null)
  const [userRoleFilter, setUserRoleFilter] = useState<'all' | '1' | '2' | '3'>('all') // 1=Usuario, 2=Supervisor, 3=Admin
  
  // Admin sidebar state
  const [showAdminSidebar, setShowAdminSidebar] = useState(false)
  const [adminSidebarTab, setAdminSidebarTab] = useState<'hojas' | 'usuarios' | 'stats' | 'reportes' | 'gps'>('hojas')
  
  // GPS Tracking states (admin)
  const [showGpsModal, setShowGpsModal] = useState(false)
  const [gpsLogs, setGpsLogs] = useState<any[]>([])
  const [gpsUsers, setGpsUsers] = useState<string[]>([])
  const [selectedGpsUser, setSelectedGpsUser] = useState<string>('')
  const [loadingGpsLogs, setLoadingGpsLogs] = useState(false)
  
  // Sheets assignment states
  const [availableSheets, setAvailableSheets] = useState<string[]>([])
  const [loadingSheets, setLoadingSheets] = useState(false)
  const [assignmentMode, setAssignmentMode] = useState<'ids' | 'sheet'>('ids')
  const [selectedSheet, setSelectedSheet] = useState<string>('')
  const [loadingSheetIds, setLoadingSheetIds] = useState(false)
  
  // Admin sheet filter state
  const [adminSelectedSheet, setAdminSelectedSheet] = useState<string>('Todos')
  
  // User sheet selection (para usuarios no admin: hoja asignada o ALTA PDV)
  const [userSelectedSheet, setUserSelectedSheet] = useState<string>('')
  
  // Pagination states
  const [currentPage, setCurrentPage] = useState(1)
  const itemsPerPage = 50
  
  // Puesto Activo/Cerrado state
  const [puestoStatus, setPuestoStatus] = useState<PuestoStatusValue | ''>('')
  
  // Autocomplete dropdown states
  const [openAutocomplete, setOpenAutocomplete] = useState<string | null>(null)
  const [autocompleteFilter, setAutocompleteFilter] = useState<{ [key: string]: string }>({})
  
  // Image upload states
  const [uploadingImage, setUploadingImage] = useState(false)
  const [imagePreview, setImagePreview] = useState<string | null>(null)
  const [uploadedImageUrl, setUploadedImageUrl] = useState<string | null>(null)
  
  // Coordenadas GPS capturadas (se guardan automáticamente, no se muestran en formulario)
  const [capturedLatitude, setCapturedLatitude] = useState<string | null>(null)
  const [capturedLongitude, setCapturedLongitude] = useState<string | null>(null)
  
  // Nuevo PDV states
  const [showNuevoPdvModal, setShowNuevoPdvModal] = useState(false)
  
  // Mobile navigation modals
  const [showMobileStats, setShowMobileStats] = useState(false)
  const [showMobileSearch, setShowMobileSearch] = useState(false)
  const [mobileSearchQuery, setMobileSearchQuery] = useState('')
  const [mobileSearchType, setMobileSearchType] = useState<'id' | 'paquete' | 'direccion'>('id')
  const [cameFromMobileSearch, setCameFromMobileSearch] = useState(false)
  const [showNuevoPdvForm, setShowNuevoPdvForm] = useState(false)
  const [savingNuevoPdv, setSavingNuevoPdv] = useState(false)
  const [nuevoPdvData, setNuevoPdvData] = useState({
    estadoKiosco: 'Abierto',
    paquete: '',
    domicilio: '',
    provincia: 'Buenos Aires',
    partido: '',
    localidad: '',
    nVendedor: '',
    distribuidora: '',
    diasAtencion: '',
    horario: '',
    escaparate: '',
    ubicacion: '',
    fachada: '',
    ventaNoEditorial: '',
    reparto: '',
    suscripciones: '',
    nombreApellido: '',
    mayorVenta: '',
    paradaOnline: '',
    telefono: '',
    correoElectronico: '',
    observaciones: '',
    comentarios: '',
    latitud: '',
    longitud: ''
  })
  const [nuevoPdvImagePreview, setNuevoPdvImagePreview] = useState<string | null>(null)
  const [nuevoPdvImageUrl, setNuevoPdvImageUrl] = useState<string | null>(null)
  const [uploadingNuevoPdvImage, setUploadingNuevoPdvImage] = useState(false)
  const [nuevoPdvFieldErrors, setNuevoPdvFieldErrors] = useState<Record<string, string>>({})
  
  // Modal para campos obligatorios faltantes
  const [showMissingFieldsModal, setShowMissingFieldsModal] = useState(false)
  const [missingFields, setMissingFields] = useState<{name: string, index: number, hint?: string}[]>([])
  const [pendingSaveAction, setPendingSaveAction] = useState<'edit' | 'nuevoPdv' | null>(null)
  
  // Popup de bienvenida / mensaje importante al iniciar
  const [showWelcomePopup, setShowWelcomePopup] = useState(false)
  const welcomePopupShownRef = useRef(false) // Para mostrar solo una vez por sesión de login

  const CLIENT_ID = '549677208908-9h6q933go4ss870pbq8gd8gaae75k338.apps.googleusercontent.com'
  const API_KEY = 'AIzaSyCJUD23abF8LcZPp7e8eiK0D5IfFoRCxUc'
  const DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
  const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email'

  const isReady = gapiInited && gisInited

  useEffect(() => {
    const gapiScript = document.createElement('script')
    gapiScript.src = 'https://apis.google.com/js/api.js'
    gapiScript.async = true
    gapiScript.defer = true
    gapiScript.onload = () => {
      if (window.gapi) {
        window.gapi.load('client', async () => {
          await window.gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: [DISCOVERY_DOC],
          })
          setGapiInited(true)
        })
      }
    }
    document.head.appendChild(gapiScript)

    const gisScript = document.createElement('script')
    gisScript.src = 'https://accounts.google.com/gsi/client'
    gisScript.async = true
    gisScript.defer = true
    gisScript.onload = () => {
      if (window.google) {
        const client = window.google.accounts.oauth2.initTokenClient({
          client_id: CLIENT_ID,
          scope: SCOPES,
          callback: '',
        })
        setTokenClient(client)
        setGisInited(true)
      }
    }
    document.head.appendChild(gisScript)

    return () => {}
  }, [])

  const handleSessionExpired = useCallback(() => {
    if (typeof sessionStorage !== 'undefined') sessionStorage.removeItem('clarin_access_token')
    document.cookie = 'clarin_session=; path=/; max-age=0'
    setAccessToken(null)
    setIsAuthorized(false)
    setUserEmail(null)
    setSheetData(null)
    setPermissions([])
    setError(null)
    setSessionExpired(true)
  }, [])

  const loadSheetData = useCallback(async (token: string, sheetName: string = '') => {
    setLoadingData(true)
    setError(null)
    
    try {
      // Si se especifica una hoja, agregarla como parámetro
      const url = sheetName ? `/api/data?sheet=${encodeURIComponent(sheetName)}` : '/api/data'
      
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      })
      
      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          setLoadingData(false)
          return
        }
        const errorData = await response.json().catch(() => ({}))
        const msg = errorData.error || 'Error cargando datos'
        const detail = errorData.details ? ` (${errorData.details})` : ''
        if (errorData.details) console.error('API /api/data:', errorData.details)
        throw new Error(msg + detail)
      }
      
      const data = await response.json()
      setSheetData(data)
      
      if (data.permissions?.isAdmin) {
        loadPermissions(token)
        loadAvailableSheets(token)
      }
    } catch (err: any) {
      setError(err.message)
    } finally {
      setLoadingData(false)
    }
  }, [handleSessionExpired])

  const { permissions: allPermissions, loadPermissions, setPermissions } = usePermissions(accessToken, handleSessionExpired)

  const loadAvailableSheets = async (token: string) => {
    setLoadingSheets(true)
    try {
      const response = await fetch('/api/sheets', {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      })
      if (response.status === 401) {
        handleSessionExpired()
        return
      }
      if (response.ok) {
        const data = await response.json()
        if (data.sheets) {
          // Filtrar hojas del sistema que no deben aparecer en el selector
          const filteredSheets = data.sheets.filter((sheet: string) => 
            sheet !== 'LOGs GPS' && 
            sheet !== 'Permisos' && 
            sheet !== 'ALTA PDV'
          )
          setAvailableSheets(filteredSheets)
        }
      }
    } catch (err) {
      console.error('Error cargando hojas:', err)
    } finally {
      setLoadingSheets(false)
    }
  }

  const loadSheetIds = async (sheetName: string): Promise<string[]> => {
    if (!accessToken) return []
    
    setLoadingSheetIds(true)
    try {
      const response = await fetch('/api/sheets', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ sheetName })
      })
      if (response.status === 401) {
        handleSessionExpired()
        return []
      }
      if (response.ok) {
        const data = await response.json()
        return data.ids || []
      }
    } catch (err) {
      console.error('Error cargando IDs de hoja:', err)
    } finally {
      setLoadingSheetIds(false)
    }
    return []
  }

  function handleAuthClick() {
    if (!tokenClient) return
    if (typeof window === 'undefined' || !window.gapi?.client || !window.google?.accounts?.oauth2) return

    setIsLoading(true)

    tokenClient.callback = async (resp: any) => {
      if (resp.error !== undefined) {
        setIsLoading(false)
        console.error('Auth error:', resp)
        return
      }
      
      const token = window.gapi.client.getToken().access_token
      setAccessToken(token)
      if (typeof sessionStorage !== 'undefined') sessionStorage.setItem('clarin_access_token', token)
      const secure = typeof window !== 'undefined' && window.location?.protocol === 'https:'
      document.cookie = `clarin_session=1; path=/; max-age=86400; samesite=lax${secure ? '; secure' : ''}`
      setIsAuthorized(true)
      setSessionExpired(false)
      setIsLoading(false)
      
      try {
        const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
          headers: { Authorization: `Bearer ${token}` }
        })
        const data = await response.json()
        setUserEmail(data.email)
      } catch (e) {
        console.error('Error getting user info:', e)
      }
      
      await loadSheetData(token, 'Todos')
    }

    if (window.gapi.client.getToken() === null) {
      tokenClient.requestAccessToken({ prompt: 'consent' })
    } else {
      tokenClient.requestAccessToken({ prompt: '' })
    }
  }

  function handleSignoutClick() {
    if (typeof window !== 'undefined' && window.gapi?.client?.getToken && window.google?.accounts?.oauth2?.revoke) {
      const token = window.gapi.client.getToken()
      if (token !== null) {
        try { window.google.accounts.oauth2.revoke(token.access_token) } catch { /* ignore */ }
        window.gapi.client.setToken('')
      }
    }
    handleSessionExpired()
  }

  // Restaurar sesión desde sessionStorage al volver desde /mapa u otra pestaña
  useEffect(() => {
    const storedToken = typeof sessionStorage !== 'undefined' ? sessionStorage.getItem('clarin_access_token') : null
    if (!storedToken) {
      setSessionChecked(true)
      return
    }
    setAccessToken(storedToken)
    setIsAuthorized(true)
    setSessionChecked(true)
    fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${storedToken}` }
    })
      .then(res => {
        if (!res.ok) throw new Error('Token inválido')
        return res.json()
      })
      .then(data => { if (data?.email) setUserEmail(data.email) })
      .catch(() => { handleSessionExpired() })
    loadSheetData(storedToken, 'Todos')
  }, [handleSessionExpired, loadSheetData])

  // Referencia para controlar el último envío de GPS (evitar spam)
  const lastGpsLogTime = useRef<number>(0)
  const GPS_LOG_COOLDOWN = 5 * 60 * 1000 // 5 minutos entre logs
  
  // Cache de ubicación para mobile (evitar pedir GPS múltiples veces)
  const cachedMobileLocation = useRef<{ latitude: number; longitude: number; timestamp: number } | null>(null)
  const LOCATION_CACHE_DURATION = 10 * 60 * 1000 // 10 minutos de validez del cache

  // Función para enviar log de GPS al servidor
  const sendGpsLog = useCallback(async (token: string, reason: string = 'login') => {
    // Solo enviar en mobile
    const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)
    
    if (!isMobile) {
      return
    }

    // Verificar cooldown (no enviar más de 1 log cada 5 minutos)
    const now = Date.now()
    if (now - lastGpsLogTime.current < GPS_LOG_COOLDOWN) {
      return
    }

    try {
      // Obtener ubicación
      if (!navigator.geolocation) {
        return
      }

      navigator.geolocation.getCurrentPosition(
        async (position) => {
          try {
            const response = await fetch('/api/log-gps', {
              method: 'POST',
              headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
              },
              body: JSON.stringify({
                latitude: position.coords.latitude,
                longitude: position.coords.longitude,
                userAgent: navigator.userAgent,
                isMobile: true,
                reason: reason // 'login', 'foreground', 'focus'
              })
            })

            if (response.ok) {
              lastGpsLogTime.current = Date.now()
              // Guardar ubicación en cache para reutilizar
              cachedMobileLocation.current = {
                latitude: position.coords.latitude,
                longitude: position.coords.longitude,
                timestamp: Date.now()
              }
            }
          } catch (error) {
            console.error('GPS Log: Error enviando datos', error)
          }
        },
        (error) => {
        },
        {
          enableHighAccuracy: false,
          timeout: 10000,
          maximumAge: 0 // Siempre obtener ubicación fresca
        }
      )
    } catch (error) {
      console.error('GPS Log: Error general', error)
    }
  }, [])

  // Efecto para enviar GPS log cuando el usuario se autentica
  useEffect(() => {
    if (isAuthorized && accessToken) {
      // Enviar log inicial después de login
      const timer = setTimeout(() => {
        sendGpsLog(accessToken, 'login')
      }, 2000)
      
      return () => clearTimeout(timer)
    }
  }, [isAuthorized, accessToken, sendGpsLog])

  // Efecto para mostrar popup de bienvenida después del login (solo una vez por sesión)
  useEffect(() => {
    if (isAuthorized && sheetData && !welcomePopupShownRef.current) {
      // Verificar si el usuario ya marcó "no mostrar más"
      const dontShowAgain = localStorage.getItem('hideWelcomePopup')
      if (!dontShowAgain) {
        // Marcar que ya mostramos el popup en esta sesión
        welcomePopupShownRef.current = true
        // Mostrar popup después de un breve delay para que cargue el dashboard
        const timer = setTimeout(() => {
          setShowWelcomePopup(true)
        }, 500)
        return () => clearTimeout(timer)
      }
    }
  }, [isAuthorized, sheetData])
  
  // Resetear el ref cuando el usuario cierra sesión
  useEffect(() => {
    if (!isAuthorized) {
      welcomePopupShownRef.current = false
    }
  }, [isAuthorized])

  // Función para cerrar popup y opcionalmente no mostrarlo más
  const handleCloseWelcomePopup = (dontShowAgain: boolean = false) => {
    if (dontShowAgain) {
      localStorage.setItem('hideWelcomePopup', 'true')
    }
    setShowWelcomePopup(false)
  }

  // Función para cargar logs de GPS (solo admin)
  const loadGpsLogs = useCallback(async (filterEmail?: string) => {
    if (!accessToken) return
    
    setLoadingGpsLogs(true)
    try {
      const url = filterEmail 
        ? `/api/gps-logs?email=${encodeURIComponent(filterEmail)}`
        : '/api/gps-logs'
      
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        }
      })
      if (response.status === 401) {
        handleSessionExpired()
        return
      }
      if (response.ok) {
        const data = await response.json()
        setGpsLogs(data.logs || [])
        setGpsUsers(data.users || [])
      }
    } catch (error) {
      console.error('Error cargando GPS logs:', error)
    } finally {
      setLoadingGpsLogs(false)
    }
  }, [accessToken, handleSessionExpired])

  // Efecto para detectar cuando la app vuelve de segundo plano
  useEffect(() => {
    if (!isAuthorized || !accessToken) return

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        // La app volvió al primer plano
        sendGpsLog(accessToken, 'foreground')
      }
    }

    const handleFocus = () => {
      // La ventana obtuvo foco
      sendGpsLog(accessToken, 'focus')
    }

    // Escuchar cambios de visibilidad (cuando vuelve de segundo plano)
    document.addEventListener('visibilitychange', handleVisibilityChange)
    // Escuchar cuando la ventana obtiene foco
    window.addEventListener('focus', handleFocus)

    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange)
      window.removeEventListener('focus', handleFocus)
    }
  }, [isAuthorized, accessToken, sendGpsLog])

  const handleEditRow = (rowIndex: number) => {
    if (!sheetData?.data?.length || !sheetData?.headers?.length) return
    if (rowIndex < 0 || rowIndex >= sheetData.data.length) return
    const rowData = [...sheetData.data[rowIndex]]
      while (rowData.length < sheetData.headers.length) {
        rowData.push('')
      }
      
      // Verificar si el PDV ya fue relevado
      const { relevadorIndex } = getAutoFillIndexes()
      const relevadorValue = relevadorIndex !== -1 ? String(rowData[relevadorIndex] || '').trim() : ''
      const isAlreadyRelevado = relevadorValue !== ''
      
      // Si ya fue relevado, pedir confirmación
      if (isAlreadyRelevado) {
        const confirmEdit = window.confirm(
          `⚠️ Este puesto ya fue relevado por: ${relevadorValue}\n\n¿Estás seguro que quieres editar este puesto ya relevado?`
        )
        if (!confirmEdit) return
      }
      
      setEditingRow(rowIndex)
      setEditedValues(rowData)
      setOriginalValues(rowData) // Guardar los valores originales del Excel
      // Preseleccionar estado según valor actual de Estado Kiosco
      const headers = sheetData.headers.map((h: string) => h.toLowerCase().trim())
      const estadoKioscoIdx = headers.findIndex((h: string) => h.includes('estado') && h.includes('kiosco'))
      const estadoVal = estadoKioscoIdx !== -1 ? String(rowData[estadoKioscoIdx] || '').trim() : ''
      if (estadoVal === 'Ahora es Cafeteria' || estadoVal === 'ahora es cafeteria') {
        setPuestoStatus('cafeteria')
      } else if (estadoVal === 'Cerrado definitivamente') {
        setPuestoStatus('cerrado')
      } else if (estadoVal === 'No se encuentra el puesto') {
        setPuestoStatus('no_encontrado')
      } else if (estadoVal === 'Zona Peligrosa') {
        setPuestoStatus('zona_peligrosa')
      } else {
        setPuestoStatus('abierto')
      }
      
      // Cargar imagen existente si hay
      const imgIndex = sheetData.headers.findIndex(h => 
        h.toLowerCase().trim() === 'img' || h.toLowerCase().trim() === 'imagen'
      )
      if (imgIndex !== -1 && rowData[imgIndex]) {
        setUploadedImageUrl(rowData[imgIndex])
        setImagePreview(rowData[imgIndex])
      } else {
        setUploadedImageUrl(null)
        setImagePreview(null)
      }
  }

  const handleCancelEdit = () => {
    const confirmCancel = window.confirm('¿Estás seguro de que deseas cancelar? Los cambios no guardados se perderán.')
    if (confirmCancel) {
      setEditingRow(null)
      setEditedValues([])
      setOriginalValues([])
      setPuestoStatus('')
      setOpenAutocomplete(null)
      setAutocompleteFilter({})
      setImagePreview(null)
      setUploadedImageUrl(null)
      setCapturedLatitude(null)
      setCapturedLongitude(null)
      
      // Si vino de la búsqueda mobile, volver a ella
      if (cameFromMobileSearch) {
        setShowMobileSearch(true)
        setCameFromMobileSearch(false)
      }
    }
  }

  // Find column indexes for auto-fill fields
  const getAutoFillIndexes = () => {
    if (!sheetData) return { fechaIndex: -1, relevadorIndex: -1 }
    
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    
    // Find "Fecha" column - buscar específicamente "fecha de relevo" primero
    let fechaIndex = headers.findIndex(h => h.includes('fecha de relevo') || h.includes('fecha relevo'))
    if (fechaIndex === -1) {
      // Si no encuentra "fecha de relevo", buscar "fecha" que NO sea otro campo
      fechaIndex = headers.findIndex(h => 
        (h.includes('fecha') && !h.includes('alta') && !h.includes('nacimiento'))
      )
    }
    
    // Find "Relevado por:" column - BÚSQUEDA MUY ESPECÍFICA
    // La columna se llama exactamente "Relevado por:" o similar
    let relevadorIndex = -1
    
    // Buscar por orden de prioridad
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i]
      // Coincidencia exacta primero
      if (h === 'relevado por:' || h === 'relevado por') {
        relevadorIndex = i
        break
      }
    }
    
    // Si no encontró exacto, buscar variaciones
    if (relevadorIndex === -1) {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i]
        // Buscar columnas que empiecen con "relevado" pero NO sean "correo"
        if ((h.startsWith('relevado') || h.startsWith('censado')) && !h.includes('correo') && !h.includes('email')) {
          relevadorIndex = i
          break
        }
      }
    }
    
    return { fechaIndex, relevadorIndex }
  }

  // Campos que se auto-rellenan cuando el puesto está cerrado
  const getCamposCerradoIndexes = () => {
    if (!sheetData) return []
    
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    const camposCerrado: number[] = []
    
    headers.forEach((h, idx) => {
      if (
        (h.includes('estado') && h.includes('kiosco')) ||
        h.includes('dias de atención') || h.includes('dias de atencion') || h === 'dias de atención' || h === 'dias de atencion' ||
        h === 'horario' || h === 'horario:' ||
        h === 'escaparate' || h === 'escaparate:' ||
        h === 'ubicacion' || h === 'ubicación' || h === 'ubicacion:' || h === 'ubicación:' ||
        (h.includes('fachada') && h.includes('puesto')) ||
        (h.includes('venta') && h.includes('no editorial')) ||
        h === 'reparto' || h === 'reparto:' ||
        h === 'suscripciones' || h === 'suscripciones:' || h === 'suscripcion' || h === 'suscripción' ||
        h === 'nombre' || h === 'nombre:' ||
        h === 'apellido' || h === 'apellido:' ||
        (h.includes('mayor') && h.includes('venta')) ||
        (h.includes('parada') && h.includes('online')) ||
        h === 'teléfono' || h === 'telefono' || h === 'teléfono:' || h === 'telefono:' ||
        h === 'correo electrónico' || h === 'correo electronico' || h === 'correo electrónico:' || h === 'correo electronico:' || h === 'email' || h === 'email:'
      ) {
        camposCerrado.push(idx)
      }
    })
    
    return camposCerrado
  }

  // Manejar cambio de estado del puesto
  const handlePuestoStatusChange = (status: PuestoStatusValue) => {
    setPuestoStatus(status)
    
    if (status === 'abierto') {
      // Restaurar los valores originales del Excel
      const camposCerradoIndexes = getCamposCerradoIndexes()
      const newValues = [...editedValues]
      
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = originalValues[idx] || ''
      })
      
      setEditedValues(newValues)
    } else if (status === 'cerrado') {
      // Solo cambiar el Estado Kiosco a "Cerrado definitivamente", el resto de campos mantienen sus valores originales
      const headers = sheetData?.headers.map(h => h.toLowerCase().trim()) || []
      const estadoKioscoIndex = headers.findIndex(h => h.includes('estado') && h.includes('kiosco'))
      
      const newValues = [...editedValues]
      
      // Restaurar todos los campos cerrados a sus valores originales primero
      const camposCerradoIndexes = getCamposCerradoIndexes()
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = originalValues[idx] || ''
      })
      
      // Luego solo cambiar el Estado Kiosco
      if (estadoKioscoIndex !== -1) {
        newValues[estadoKioscoIndex] = 'Cerrado definitivamente'
      }
      
      setEditedValues(newValues)
    } else if (status === 'no_encontrado') {
      // Solo cambiar el Estado Kiosco a "No se encuentra el puesto", el resto de campos mantienen sus valores originales
      const headers = sheetData?.headers.map(h => h.toLowerCase().trim()) || []
      const estadoKioscoIndex = headers.findIndex(h => h.includes('estado') && h.includes('kiosco'))
      
      const newValues = [...editedValues]
      
      // Restaurar todos los campos cerrados a sus valores originales primero
      const camposCerradoIndexes = getCamposCerradoIndexes()
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = originalValues[idx] || ''
      })
      
      // Luego solo cambiar el Estado Kiosco
      if (estadoKioscoIndex !== -1) {
        newValues[estadoKioscoIndex] = 'No se encuentra el puesto'
      }
      
      setEditedValues(newValues)
    } else if (status === 'zona_peligrosa') {
      // Solo cambiar el Estado Kiosco a "Zona Peligrosa", el resto de campos mantienen sus valores originales
      const headers = sheetData?.headers.map(h => h.toLowerCase().trim()) || []
      const estadoKioscoIndex = headers.findIndex(h => h.includes('estado') && h.includes('kiosco'))
      
      const newValues = [...editedValues]
      
      // Restaurar todos los campos cerrados a sus valores originales primero
      const camposCerradoIndexes = getCamposCerradoIndexes()
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = originalValues[idx] || ''
      })
      
      // Luego solo cambiar el Estado Kiosco a Zona Peligrosa
      if (estadoKioscoIndex !== -1) {
        newValues[estadoKioscoIndex] = 'Zona Peligrosa'
      }
      
      setEditedValues(newValues)
    } else if (status === 'cafeteria') {
      // Solo cambiar el Estado Kiosco a "Ahora es Cafeteria", el resto de campos mantienen sus valores
      const headers = sheetData?.headers.map(h => h.toLowerCase().trim()) || []
      const estadoKioscoIndex = headers.findIndex(h => h.includes('estado') && h.includes('kiosco'))
      
      const newValues = [...editedValues]
      
      if (estadoKioscoIndex !== -1) {
        newValues[estadoKioscoIndex] = 'Ahora es Cafeteria'
      }
      
      setEditedValues(newValues)
    }
  }

  // Función para obtener la ubicación GPS actual
  const getCurrentLocation = (): Promise<{latitude: number, longitude: number, error?: string} | null> => {
    return new Promise((resolve) => {
      if (!navigator.geolocation) {
        alert('⚠️ Tu navegador no soporta geolocalización.\nIntenta con otro navegador.')
        resolve(null)
        return
      }

      navigator.geolocation.getCurrentPosition(
        (position) => {
          resolve({
            latitude: position.coords.latitude,
            longitude: position.coords.longitude
          })
        },
        (error) => {
          let errorMessage = ''
          switch (error.code) {
            case error.PERMISSION_DENIED:
              errorMessage = '⚠️ Permiso de ubicación denegado.\n\nPor favor, permite el acceso a la ubicación en la configuración de tu navegador.'
              break
            case error.POSITION_UNAVAILABLE:
              errorMessage = '⚠️ Ubicación no disponible.\n\nAsegúrate de tener el GPS activado y estar en un lugar con buena señal.'
              break
            case error.TIMEOUT:
              errorMessage = '⚠️ Tiempo de espera agotado.\n\nNo se pudo obtener la ubicación. Intenta de nuevo en un lugar con mejor señal GPS.'
              break
            default:
              errorMessage = '⚠️ Error desconocido al obtener ubicación.\n\nIntenta de nuevo.'
          }
          alert(errorMessage)
          resolve(null)
        },
        {
          enableHighAccuracy: true,
          timeout: 15000, // Aumentado a 15 segundos
          maximumAge: 0
        }
      )
    })
  }

  // Función para subir imagen a ImgBB
  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>, captureLocation: boolean = false) => {
    const file = e.target.files?.[0]
    if (!file || !accessToken) return

    // Limpiar el input file inmediatamente para evitar re-disparos
    const inputElement = e.target
    
    // Validar que sea una imagen
    if (!file.type.startsWith('image/')) {
      alert('Por favor selecciona un archivo de imagen válido')
      inputElement.value = ''
      return
    }

    // Validar tamaño (máx 32MB para ImgBB)
    if (file.size > 32 * 1024 * 1024) {
      alert('La imagen es demasiado grande. Máximo 32MB')
      inputElement.value = ''
      return
    }

    // Mostrar preview
    const reader = new FileReader()
    reader.onload = (event) => {
      setImagePreview(event.target?.result as string)
    }
    reader.readAsDataURL(file)
    
    // Limpiar el input para evitar problemas de re-render
    inputElement.value = ''

    // Subir a ImgBB
    setUploadingImage(true)
    try {
      // Si se solicita capturar ubicación (cámara), verificar si ya existen coordenadas
      let location: {latitude: number, longitude: number} | null = null
      let shouldCaptureLocation = captureLocation
      
      if (captureLocation && sheetData && editingRow !== null) {
        const headers = sheetData.headers.map(h => h.toLowerCase().trim())
        const latIndex = headers.findIndex(h => h === 'latitud' || h === 'lat')
        const lngIndex = headers.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')
        
        if (latIndex !== -1 && lngIndex !== -1) {
          const existingLat = String(sheetData.data[editingRow][latIndex] || '').trim()
          const existingLng = String(sheetData.data[editingRow][lngIndex] || '').trim()
          
          if (existingLat || existingLng) {
            shouldCaptureLocation = window.confirm(
              `📍 Este registro ya tiene coordenadas:\n\n` +
              `Latitud: ${existingLat || '(vacío)'}\n` +
              `Longitud: ${existingLng || '(vacío)'}\n\n` +
              `¿Deseas actualizar las coordenadas con tu ubicación actual?\n\n` +
              `(Presiona "Cancelar" para subir solo la foto sin cambiar las coordenadas)`
            )
          }
        }
      }
      
      if (shouldCaptureLocation) {
        location = await getCurrentLocation()
      }

      const formData = new FormData()
      formData.append('image', file)

      const response = await fetch('/api/upload', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        body: formData
      })

      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          return
        }
        let errorMessage = 'Error al subir imagen'
        try {
          const errorData = await response.json()
          errorMessage = errorData.error || errorMessage
        } catch {
          if (response.status === 413) {
            errorMessage = 'La imagen es demasiado grande. Máximo 32MB.'
          } else if (response.status >= 500) {
            errorMessage = 'Error del servidor. Intenta de nuevo en unos minutos.'
          }
        }
        throw new Error(errorMessage)
      }

      const data = await response.json()
      if (!data.imageUrl) {
        throw new Error('El servidor no devolvió la URL de la imagen. Intenta de nuevo.')
      }
      setUploadedImageUrl(data.imageUrl)

      // Actualizar solo el campo IMG en editedValues
      if (sheetData) {
        const headers = sheetData.headers.map(h => h.toLowerCase().trim())
        const imgIndex = headers.findIndex(h => h === 'img' || h === 'imagen')
        
        if (imgIndex !== -1) {
          setEditedValues(prevValues => {
            const newValues = [...prevValues]
            // Asegurar longitud correcta
            while (newValues.length < sheetData.headers.length) {
              newValues.push('')
            }
            newValues[imgIndex] = data.imageUrl
            return newValues
          })
        }
      }

      // Guardar coordenadas en estados separados (se agregarán al guardar)
      if (location) {
        setCapturedLatitude(location.latitude.toFixed(6))
        setCapturedLongitude(location.longitude.toFixed(6))
      }

      // Mensaje de éxito
      let successMessage = 'Imagen subida correctamente'
      if (location) {
        successMessage = `Imagen subida correctamente. Ubicación: ${location.latitude.toFixed(6)}, ${location.longitude.toFixed(6)}`
      } else if (captureLocation && !shouldCaptureLocation) {
        successMessage = 'Imagen subida correctamente. Coordenadas anteriores conservadas.'
      } else if (captureLocation) {
        successMessage = 'Imagen subida correctamente. No se pudo obtener la ubicación GPS.'
      }
      toast.success(successMessage)
    } catch (error: any) {
      console.error('Error uploading image:', error)
      toast.error('Error al subir la imagen: ' + error.message)
      setImagePreview(null)
    } finally {
      setUploadingImage(false)
    }
  }

  // Limpiar imagen y coordenadas
  const handleClearImage = () => {
    setImagePreview(null)
    setUploadedImageUrl(null)
    setCapturedLatitude(null)
    setCapturedLongitude(null)
    if (sheetData) {
      const imgIndex = sheetData.headers.findIndex(h => 
        h.toLowerCase().trim() === 'img' || h.toLowerCase().trim() === 'imagen'
      )
      if (imgIndex !== -1) {
        const newValues = [...editedValues]
        // Asegurar longitud correcta
        while (newValues.length < sheetData.headers.length) {
          newValues.push('')
        }
        newValues[imgIndex] = ''
        setEditedValues(newValues)
      }
    }
  }

  // Estado para indicar qué fila está guardando ubicación
  const [savingLocationRow, setSavingLocationRow] = useState<number | null>(null)
  const [locationModalRow, setLocationModalRow] = useState<number | null>(null)
  const [showMapPicker, setShowMapPicker] = useState(false)
  const [manualCoords, setManualCoords] = useState<{ lat: number; lng: number } | null>(null)
  const [savingManualLocation, setSavingManualLocation] = useState(false)
  const [addressSearch, setAddressSearch] = useState('')
  const [searchingAddress, setSearchingAddress] = useState(false)
  const [showLocationWarning, setShowLocationWarning] = useState(false)

  // Cerrar modales con Escape (accesibilidad) — después de declarar showMapPicker, locationModalRow, etc.
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      if (showMapPicker) setShowMapPicker(false)
      else if (locationModalRow !== null) setLocationModalRow(null)
      else if (editingRow !== null) {
        setEditingRow(null)
        setEditedValues([])
        setOriginalValues([])
      }
      else if (showPermissionsPanel) setShowPermissionsPanel(false)
      else if (showStats) setShowStats(false)
      else if (showAdminSidebar) setShowAdminSidebar(false)
      else if (showMissingFieldsModal) setShowMissingFieldsModal(false)
      else if (showNuevoPdvModal) setShowNuevoPdvModal(false)
      else if (showGpsModal) setShowGpsModal(false)
      else if (showWelcomePopup) setShowWelcomePopup(false)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [showMapPicker, locationModalRow, editingRow, showPermissionsPanel, showStats, showAdminSidebar, showMissingFieldsModal, showNuevoPdvModal, showGpsModal, showWelcomePopup])

  // Función para abrir el modal de opciones de ubicación
  const handleLocationClick = (rowIndex: number) => {
    setLocationModalRow(rowIndex)
  }

  // Función para guardar ubicación usando GPS actual
  const handleSaveCurrentLocation = async () => {
    if (locationModalRow === null) return
    
    const rowToSave = locationModalRow
    
    // Siempre pedir GPS fresco (ubicación actual, no la del cache)
    const confirmed = window.confirm(
      '📍 ¿Guardar tu ubicación actual?\n\n' +
      'Se obtendrán las coordenadas GPS de tu dispositivo en este momento.'
    )
    if (!confirmed) return
    
    // No cerrar el modal hasta que termine - mostrar animación
    await handleSaveLocation(rowToSave, 'gps')
    setLocationModalRow(null)
    setCameFromMobileSearch(false) // No volver a búsqueda después de guardar
  }

  // Función para buscar dirección y obtener coordenadas
  const handleSearchAddress = async () => {
    if (!addressSearch.trim()) {
      alert('⚠️ Ingresa una dirección para buscar')
      return
    }

    setSearchingAddress(true)
    try {
      // Usar API local para evitar problemas de CORS en mobile
      const response = await fetch(
        `/api/geocode?q=${encodeURIComponent(addressSearch)}`
      )

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}))
        console.error('Geocode API error:', response.status, errorData)
        throw new Error(errorData.error || 'Error en la búsqueda')
      }

      const data = await response.json()
      const results = data.results
      
      if (!results || results.length === 0) {
        alert('📍 No se encontró la dirección.\n\nIntenta con otro formato:\n- "Santa Fe 300, Palermo"\n- "Av. Corrientes 1234"\n- "Florida y Lavalle"')
        return
      }

      const { lat, lon, display_name } = results[0]
      setManualCoords({ lat: parseFloat(lat), lng: parseFloat(lon) })
      
      // Mostrar dirección encontrada
      alert(`✅ Ubicación encontrada:\n\n${display_name}\n\nAhora puedes ajustar el pin si es necesario.`)
    } catch (error) {
      alert('❌ Error buscando dirección.\n\nIntenta de nuevo o ingresa las coordenadas manualmente.')
    } finally {
      setSearchingAddress(false)
    }
  }

  // Función para abrir el selector de mapa
  const handleOpenMapPicker = async () => {
    if (locationModalRow === null || !sheetData?.data?.length || !sheetData?.headers?.length) return
    if (locationModalRow < 0 || locationModalRow >= sheetData.data.length) return
    const row = sheetData.data[locationModalRow]
    if (!row) return
    
    // Obtener coordenadas existentes si las hay para centrar el mapa
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    const latIndex = headers.findIndex(h => h === 'latitud' || h === 'lat')
    const lngIndex = headers.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')
    
    let initialLat = -34.6037 // Buenos Aires por defecto
    let initialLng = -58.3816
    
    if (latIndex !== -1 && lngIndex !== -1) {
      const existingLat = parseFloat(String(row[latIndex] || ''))
      const existingLng = parseFloat(String(row[lngIndex] || ''))
      if (!isNaN(existingLat) && !isNaN(existingLng)) {
        initialLat = existingLat
        initialLng = existingLng
      }
    }
    
    // Pre-cargar el domicilio del puesto en el campo de búsqueda
    const domicilioIndex = headers.findIndex(h => h === 'domicilio' || h.includes('direccion') || h.includes('dirección'))
    const localidadIndex = headers.findIndex(h => h === 'localidad' || h === 'ciudad' || h === 'partido')
    
    let addressSuggestion = ''
    if (domicilioIndex !== -1) {
      const domicilio = String(row[domicilioIndex] || '').trim()
      if (domicilio) {
        addressSuggestion = domicilio
        if (localidadIndex !== -1) {
          const localidad = String(row[localidadIndex] || '').trim()
          if (localidad) addressSuggestion += `, ${localidad}`
        }
      }
    }
    
    setManualCoords({ lat: initialLat, lng: initialLng })
    setAddressSearch(addressSuggestion)
    setShowMapPicker(true)
    
    // Si hay domicilio, buscar automáticamente y mostrar aviso
    if (addressSuggestion) {
      setSearchingAddress(true)
      setShowLocationWarning(true) // Mostrar popup de aviso
      
      try {
        const response = await fetch(
          `/api/geocode?q=${encodeURIComponent(addressSuggestion)}`
        )
        
        if (response.ok) {
          const data = await response.json()
          const results = data.results
          
          if (results && results.length > 0) {
            const { lat, lon } = results[0]
            setManualCoords({ lat: parseFloat(lat), lng: parseFloat(lon) })
          }
        }
      } catch (error) {
        console.error('Error buscando dirección automáticamente:', error)
      } finally {
        setSearchingAddress(false)
      }
    }
  }

  // Función para guardar ubicación manual del mapa
  const handleSaveManualLocation = async () => {
    if (locationModalRow === null) {
      alert('⚠️ Error interno: No hay fila seleccionada.\n\nCierra el modal y vuelve a intentar.')
      return
    }
    if (!manualCoords) {
      alert('⚠️ No hay coordenadas seleccionadas.\n\nSelecciona una ubicación en el mapa.')
      return
    }
    if (isNaN(manualCoords.lat) || isNaN(manualCoords.lng)) {
      alert('⚠️ Coordenadas inválidas.\n\nSelecciona una ubicación válida en el mapa.')
      return
    }
    if (manualCoords.lat < -90 || manualCoords.lat > 90 || manualCoords.lng < -180 || manualCoords.lng > 180) {
      alert('⚠️ Coordenadas fuera de rango.\n\nLatitud debe estar entre -90 y 90.\nLongitud debe estar entre -180 y 180.')
      return
    }
    
    setSavingManualLocation(true)
    try {
      await handleSaveLocation(locationModalRow, 'manual', manualCoords)
      setShowMapPicker(false)
      setLocationModalRow(null)
      setManualCoords(null)
      setShowLocationWarning(false)
      setCameFromMobileSearch(false) // No volver a búsqueda después de guardar
    } finally {
      setSavingManualLocation(false)
    }
  }

  // Función para guardar solo la ubicación GPS de una fila
  const handleSaveLocation = async (rowIndex: number, mode: 'gps' | 'manual' = 'gps', coords?: { lat: number; lng: number }) => {
    if (!accessToken) {
      alert('⚠️ Sesión expirada.\n\nPor favor, cierra sesión y vuelve a ingresar.')
      return
    }
    if (!sheetData?.data?.length || !sheetData?.headers?.length) {
      alert('⚠️ No hay datos cargados.\n\nRecarga la página e intenta de nuevo.')
      return
    }
    if (rowIndex < 0 || rowIndex >= sheetData.data.length) return

    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    const latIndex = headers.findIndex(h => h === 'latitud' || h === 'lat')
    const lngIndex = headers.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')

    if (latIndex === -1 || lngIndex === -1) {
      alert('⚠️ Las columnas "latitud" y/o "longitud" no existen en el Excel.\nAgrega estas columnas para poder guardar la ubicación.')
      return
    }

    setSavingLocationRow(rowIndex)
    try {
      let latitude: number
      let longitude: number

      if (mode === 'manual' && coords) {
        // Usar coordenadas manuales del mapa
        latitude = coords.lat
        longitude = coords.lng
      } else {
        // Obtener ubicación GPS actual
        const location = await getCurrentLocation()
        if (!location) {
          alert('⚠️ No se pudo obtener la ubicación GPS.\nAsegúrate de permitir el acceso a la ubicación.')
          setSavingLocationRow(null) // Resetear estado antes de salir
          return
        }
        latitude = location.latitude
        longitude = location.longitude
      }

      const rowId = sheetData.data[rowIndex][0]

      // Preparar los valores a guardar (copiar fila actual y actualizar coordenadas)
      // Asegurar que el array tenga la misma longitud que los headers
      const valuesToSave = [...sheetData.data[rowIndex]]
      while (valuesToSave.length < sheetData.headers.length) {
        valuesToSave.push('')
      }
      valuesToSave[latIndex] = latitude.toFixed(6)
      valuesToSave[lngIndex] = longitude.toFixed(6)
      
      // Guardar dispositivo usado
      const dispositivoIndex = headers.findIndex(h => 
        h === 'dispositivo' || h.includes('dispositivo') || h === 'device'
      )
      if (dispositivoIndex !== -1) {
        const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)
        valuesToSave[dispositivoIndex] = isMobile ? 'Mobile' : 'PC'
      }

      const response = await fetch('/api/update', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          rowId: rowId,
          values: valuesToSave,
          sheetName: adminSelectedSheet && adminSelectedSheet !== 'Todos' ? adminSelectedSheet : undefined
        })
      })

      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          return
        }
        let errorMessage = 'Error guardando ubicación'
        try {
          const errorData = await response.json()
          errorMessage = errorData.error || errorMessage
        } catch {
          if (response.status === 403) {
            errorMessage = 'No tienes permiso para editar este registro.'
          } else if (response.status === 404) {
            errorMessage = 'Registro no encontrado. Recarga la página.'
          } else if (response.status >= 500) {
            errorMessage = 'Error del servidor. Intenta de nuevo en unos minutos.'
          }
        }
        throw new Error(errorMessage)
      }

      // Actualizar datos locales
      const newData = [...sheetData.data]
      newData[rowIndex] = valuesToSave
      setSheetData({ ...sheetData, data: newData })

      const modeText = mode === 'manual' ? '(manual)' : '(GPS)'
      toast.success(`Ubicación guardada ${modeText}. ${latitude.toFixed(6)}, ${longitude.toFixed(6)}`)
    } catch (err: any) {
      const errorMsg = err.message || 'Error desconocido'
      toast.error(`Error al guardar ubicación: ${errorMsg}`)
    } finally {
      setSavingLocationRow(null)
    }
  }

  const handleSaveRow = async (skipValidation: boolean = false) => {
    if (!accessToken || editingRow === null || !sheetData?.data?.length || !sheetData?.headers?.length || !userEmail) return
    if (editingRow < 0 || editingRow >= sheetData.data.length) return
    
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    
    if (!skipValidation) {
      const camposFaltantes: {name: string, index: number, hint?: string}[] = []
      
      // Paquete es obligatorio SIEMPRE (en todos los estados)
      const paqueteIndex = headers.findIndex(h => h.includes('paquete'))
      if (paqueteIndex !== -1 && !String(editedValues[paqueteIndex] || '').trim()) {
        camposFaltantes.push({ name: 'Paquete', index: paqueteIndex })
      }
      
      // Validar campos obligatorios adicionales solo si el puesto está activo o es cafeteria
      if (puestoStatus === 'abierto' || puestoStatus === 'cafeteria') {
        // Buscar índice de Venta productos no editoriales
        const ventaNoEditorialIndex = headers.findIndex(h => 
          h.includes('venta') && h.includes('no editorial')
        )
        
        // Buscar índice de Teléfono
        const telefonoIndex = headers.findIndex(h => 
          h.includes('telefono') || h.includes('teléfono')
        )
        
        if (ventaNoEditorialIndex !== -1 && !String(editedValues[ventaNoEditorialIndex] || '').trim()) {
          camposFaltantes.push({ name: 'Venta productos no editoriales', index: ventaNoEditorialIndex })
        }
        
        if (telefonoIndex !== -1 && !String(editedValues[telefonoIndex] || '').trim()) {
          camposFaltantes.push({ name: 'Teléfono', index: telefonoIndex, hint: 'Poner 0 si no se obtiene' })
        }
      }
      
      if (camposFaltantes.length > 0) {
        setMissingFields(camposFaltantes)
        setPendingSaveAction('edit')
        setShowMissingFieldsModal(true)
        return
      }
    }
    
    const confirmSave = window.confirm('¿Estás seguro de que deseas guardar los cambios?')
    if (!confirmSave) return
    
    setSaving(true)
    try {
      const rowId = sheetData.data[editingRow][0]
      const { fechaIndex, relevadorIndex } = getAutoFillIndexes()
      
      // Auto-fill date and relevador fields
      // IMPORTANTE: Asegurar que el array tenga la misma longitud que los headers
      // para evitar desplazamiento de columnas
      const valuesToSave = [...editedValues]
      while (valuesToSave.length < sheetData.headers.length) {
        valuesToSave.push('')
      }
      
      // Set current date in fecha field
      if (fechaIndex !== -1) {
        const today = new Date()
        const dateStr = today.toLocaleDateString('es-AR', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })
        valuesToSave[fechaIndex] = dateStr
      }
      
      // Relevado por: solo lo llenamos con el email del usuario si es usuario normal.
      // Admin y supervisor no deben pisar esta columna para no superponer el correo del relevador original.
      const isAdminOrSupervisor = sheetData.permissions?.role === 'admin' || sheetData.permissions?.role === 'supervisor'
      if (relevadorIndex !== -1) {
        if (isAdminOrSupervisor) {
          const originalRelevador = sheetData.data[editingRow]?.[relevadorIndex]
          valuesToSave[relevadorIndex] = originalRelevador !== undefined && originalRelevador !== null ? originalRelevador : ''
        } else {
          valuesToSave[relevadorIndex] = userEmail
        }
      }
      
      // Detectar y guardar dispositivo
      const dispositivoIndex = headers.findIndex(h => 
        h === 'dispositivo' || h.includes('dispositivo') || h === 'device'
      )
      if (dispositivoIndex !== -1) {
        const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)
        valuesToSave[dispositivoIndex] = isMobile ? 'Mobile' : 'PC'
      }
      
      // Agregar coordenadas GPS si fueron capturadas
      if (capturedLatitude || capturedLongitude) {
        const latIndex = headers.findIndex(h => h === 'latitud' || h === 'lat')
        const lngIndex = headers.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')
        
        if (latIndex !== -1 && capturedLatitude) {
          valuesToSave[latIndex] = capturedLatitude
        }
        if (lngIndex !== -1 && capturedLongitude) {
          valuesToSave[lngIndex] = capturedLongitude
        }
      }
      
      // Normalizar Localidad/Barrio y Partido a Title Case para consistencia
      const localidadIdx = getLocalidadIndex()
      const partidoIdx = getPartidoIndex()
      
      if (localidadIdx !== -1 && valuesToSave[localidadIdx]) {
        valuesToSave[localidadIdx] = toTitleCase(String(valuesToSave[localidadIdx]).trim())
      }
      if (partidoIdx !== -1 && valuesToSave[partidoIdx]) {
        valuesToSave[partidoIdx] = toTitleCase(String(valuesToSave[partidoIdx]).trim())
      }
      
      const response = await fetch('/api/update', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          rowId: rowId,
          values: valuesToSave,
          sheetName: adminSelectedSheet && adminSelectedSheet !== 'Todos' ? adminSelectedSheet : undefined
        })
      })
      
      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          return
        }
        const errorData = await response.json().catch(() => ({}))
        throw new Error(errorData.error || 'Error guardando datos')
      }
      
      const newData = [...sheetData.data]
      newData[editingRow] = valuesToSave
      setSheetData({ ...sheetData, data: newData })
      setEditingRow(null)
      setEditedValues([])
      setOpenAutocomplete(null)
      setAutocompleteFilter({})
      setImagePreview(null)
      setUploadedImageUrl(null)
      setCapturedLatitude(null)
      setCapturedLongitude(null)
      // Resetear flag de búsqueda (no volver a la búsqueda después de guardar)
      setCameFromMobileSearch(false)
    } catch (err: any) {
      toast.error('Error guardando datos: ' + err.message)
    } finally {
      setSaving(false)
    }
  }

  const handleSavePermissions = async (email: string, ids: string[], sheetName: string = '', level?: UserLevel) => {
    if (!accessToken) return
    
    setSavingPermissions(true)
    try {
      const response = await fetch('/api/permissions', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          email: email.trim(),
          allowedIds: ids,
          assignedSheet: sheetName,
          level: level
        })
      })
      
      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          return
        }
        const errorData = await response.json().catch(() => ({}))
        throw new Error(errorData.error || 'Error guardando permisos')
      }
      
      await loadPermissions(accessToken)
      setEditingPermission(null)
      setNewPermIds('')
      setNewUserEmail('')
    } catch (err: any) {
      toast.error('Error guardando permisos: ' + err.message)
    } finally {
      setSavingPermissions(false)
    }
  }

  // Considera que el valor de celda "Relevado por" coincide con el email del usuario (varios formatos)
  const relevadoPorMatchesUser = (cellValue: string, userEmail: string): boolean => {
    const cell = String(cellValue || '').toLowerCase().trim()
    const user = String(userEmail || '').toLowerCase().trim()
    if (!cell || !user) return false
    if (cell === user) return true
    const userLocal = user.split('@')[0]
    if (cell === userLocal) return true
    if (user.startsWith(cell + '@')) return true
    if (cell.includes('@') && cell === user) return true
    return false
  }

  // Calculate stats for a user based on "Relevador por:" field
  const getUserStats = (userEmailPerm: string, allowedIds: string[]) => {
    if (!sheetData) return { relevados: 0, faltantes: 0, total: 0 }
    
    const totalAssigned = allowedIds.length
    
    // Find the "Relevador por:" column index (mismo criterio que getAutoFillIndexes)
    const headers = sheetData.headers.map(h => String(h || '').toLowerCase().trim())
    let relevadorIndex = headers.findIndex(h => 
      h === 'relevado por:' || h === 'relevado por'
    )
    if (relevadorIndex === -1) {
      relevadorIndex = headers.findIndex(h => 
        (h.startsWith('relevado') || h.startsWith('censado')) && !h.includes('correo') && !h.includes('email')
      )
    }
    if (relevadorIndex === -1) {
      relevadorIndex = headers.findIndex(h => 
        h.includes('relevador') || h.includes('relevado por') || h.includes('censado por')
      )
    }
    
    if (relevadorIndex === -1) {
      const existingIds = sheetData.data.map(row => String(row[0] || '').toLowerCase())
      const relevados = allowedIds.filter(id => existingIds.includes(id.toLowerCase())).length
      return { relevados, faltantes: totalAssigned - relevados, total: totalAssigned }
    }
    
    const allowedIdsLower = allowedIds.map(id => id.toLowerCase())
    let relevados = 0
    for (const row of sheetData.data) {
      const rowId = String(row[0] || '').toLowerCase()
      const relevadorCell = String(row[relevadorIndex] || '').trim()
      
      if (!allowedIdsLower.includes(rowId)) continue
      if (relevadoPorMatchesUser(relevadorCell, userEmailPerm)) relevados++
    }
    
    return { relevados, faltantes: totalAssigned - relevados, total: totalAssigned }
  }

  // Get Paquete column index
  const getPaqueteIndex = () => {
    if (!sheetData) return -1
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    return headers.findIndex(h => h.includes('paquete'))
  }

  // Get Domicilio/Dirección column index
  const getDomicilioIndex = () => {
    if (!sheetData) return -1
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    return headers.findIndex(h => h.includes('domicilio') || h === 'direccion' || h === 'dirección')
  }

  // Get Partido column index
  const getPartidoIndex = () => {
    if (!sheetData) return -1
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    return headers.findIndex(h => h === 'partido' || h === 'partido:')
  }

  // Get Localidad column index
  const getLocalidadIndex = () => {
    if (!sheetData) return -1
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    return headers.findIndex(h => 
      h === 'localidad' || h === 'localidad:' || 
      h === 'localidad/barrio' || h === 'localidad/barrio:' ||
      h === 'localidad / barrio' || h === 'localidad / barrio:' ||
      h === 'barrio' || h === 'barrio:'
    )
  }

  // Índices de columnas Latitud / Longitud (para filtros censados mapeados / sin mapear)
  const getLatLngIndexes = () => {
    if (!sheetData) return { latIndex: -1, lngIndex: -1 }
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    const latIndex = headers.findIndex(h => h === 'latitud' || h === 'lat')
    const lngIndex = headers.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')
    return { latIndex, lngIndex }
  }

  // Normalizar texto a Title Case (primera letra de cada palabra en mayúscula)
  const toTitleCase = (str: string): string => {
    return str
      .toLowerCase()
      .split(' ')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ')
  }

  // Obtener valores únicos de una columna para los filtros (normalizados)
  const getUniqueColumnValues = (columnIndex: number): string[] => {
    if (!sheetData || columnIndex === -1) return []
    const valuesMap = new Map<string, string>() // lowercase -> titleCase
    sheetData.data.forEach(row => {
      const value = String(row[columnIndex] || '').trim()
      if (value) {
        const lowerKey = value.toLowerCase()
        // Solo guardar si no existe, para mantener consistencia
        if (!valuesMap.has(lowerKey)) {
          valuesMap.set(lowerKey, toTitleCase(value))
        }
      }
    })
    return Array.from(valuesMap.values()).sort()
  }

  // Valores únicos para los filtros de ubicación
  const partidoOptions = getUniqueColumnValues(getPartidoIndex())
  const localidadOptions = filterPartido 
    ? getUniqueColumnValues(getLocalidadIndex()).filter(loc => {
        const partidoIdx = getPartidoIndex()
        const localidadIdx = getLocalidadIndex()
        if (partidoIdx === -1 || localidadIdx === -1) return true
        // Comparación case-insensitive
        return sheetData?.data.some(row => 
          String(row[partidoIdx] || '').trim().toLowerCase() === filterPartido.toLowerCase() &&
          String(row[localidadIdx] || '').trim().toLowerCase() === loc.toLowerCase()
        )
      })
    : getUniqueColumnValues(getLocalidadIndex())

  // Campos para estadísticas con gráficos
  const statsFields = [
    'estado kiosco',
    'dias de atención',
    'dias de atencion', 
    'horario',
    'escaparate',
    'ubicación',
    'ubicacion',
    'fachada puesto',
    'fachada de puesto',
    'venta productos no editoriales',
    'venta de productos no editoriales',
    'suscripciones',
    'mayor venta',
    'utiliza parada online',
    'utiliza parada online?',
    'reparto'
  ]

  // Calcular estadísticas para gráficos (solo PDV relevados)
  const getFieldStats = () => {
    if (!sheetData) return []
    
    const headers = sheetData.headers
    const headersLower = headers.map(h => h.toLowerCase().trim())
    const results: { fieldName: string; data: { label: string; count: number; color: string }[] }[] = []
    
    // Obtener índice de la columna "Relevado por:"
    const { relevadorIndex } = getAutoFillIndexes()
    
    // Filtrar solo las filas que ya fueron relevadas
    const relevadosData = sheetData.data.filter(row => {
      if (relevadorIndex === -1) return false
      const relevadorValue = String(row[relevadorIndex] || '').trim()
      return relevadorValue !== ''
    })
    
    // Colores para los gráficos
    const colors = [
      '#22C55E', '#3B82F6', '#F59E0B', '#EF4444', '#8B5CF6', 
      '#EC4899', '#06B6D4', '#84CC16', '#F97316', '#6366F1',
      '#14B8A6', '#A855F7', '#FB7185', '#FBBF24', '#34D399'
    ]
    
    headersLower.forEach((header, idx) => {
      // Verificar si este campo es uno de los que queremos graficar
      const isStatsField = statsFields.some(field => 
        header.includes(field) || field.includes(header)
      )
      
      if (isStatsField) {
        const counts: { [key: string]: number } = {}
        
        // Contar valores solo de PDV relevados
        relevadosData.forEach(row => {
          const value = String(row[idx] || '').trim()
          if (value) {
            counts[value] = (counts[value] || 0) + 1
          }
        })
        
        // Convertir a array y ordenar por cantidad
        const data = Object.entries(counts)
          .map(([label, count], i) => ({
            label,
            count,
            color: colors[i % colors.length]
          }))
          .sort((a, b) => b.count - a.count)
        
        if (data.length > 0) {
          results.push({
            fieldName: headers[idx],
            data
          })
        }
      }
    })
    
    return results
  }

  // Cargar estadísticas de todas las hojas (solo admin)
  const loadAllSheetsStats = async () => {
    if (!accessToken || !sheetData?.permissions?.isAdmin) return
    
    setLoadingStats(true)
    try {
      const response = await fetch('/api/stats', {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        }
      })
      if (response.status === 401) {
        handleSessionExpired()
        return
      }
      if (response.ok) {
        const data = await response.json()
        setAllSheetsStats(data.sheets || {})
      }
    } catch (error) {
      console.error('Error cargando estadísticas:', error)
    } finally {
      setLoadingStats(false)
    }
  }

  // Calcular estadísticas para una hoja específica o el total
  const getFieldStatsForSheet = (sheetName: string) => {
    let dataToAnalyze: any[][] = []
    let headersToUse: string[] = []
    
    if (sheetName === 'Total') {
      // Combinar todas las hojas excepto "Hoja 1"
      const sheetNames = Object.keys(allSheetsStats).filter(name => name !== 'Hoja 1')
      
      if (sheetNames.length > 0) {
        // Usar los headers de la primera hoja disponible
        const firstSheet = sheetNames[0]
        const firstData = allSheetsStats[firstSheet]
        if (firstData && firstData.length > 0) {
          headersToUse = firstData[0] as string[]
          
          // Combinar datos de todas las hojas (sin headers)
          sheetNames.forEach(name => {
            const sheetData = allSheetsStats[name]
            if (sheetData && sheetData.length > 1) {
              dataToAnalyze = dataToAnalyze.concat(sheetData.slice(1))
            }
          })
        }
      }
    } else {
      // Datos de una hoja específica
      const sheetData = allSheetsStats[sheetName]
      if (sheetData && sheetData.length > 0) {
        headersToUse = sheetData[0] as string[]
        dataToAnalyze = sheetData.slice(1)
      }
    }
    
    if (headersToUse.length === 0 || dataToAnalyze.length === 0) {
      return { stats: [], total: 0, relevados: 0, pendientes: 0, relevadosConCoordenadas: 0 }
    }
    
    const headersLower = headersToUse.map(h => String(h).toLowerCase().trim())
    const results: { fieldName: string; data: { label: string; count: number; color: string }[] }[] = []
    
    const relevadorIndex = headersLower.findIndex(h => 
      h.includes('relevador') || h.includes('relevado por') || h.includes('censado por')
    )
    const latIndex = headersLower.findIndex(h => h === 'latitud' || h === 'lat')
    const lngIndex = headersLower.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')
    
    const parseCoord = (val: unknown): number => {
      const s = String(val ?? '').trim().replace(',', '.')
      if (!s) return NaN
      const n = parseFloat(s)
      return typeof n === 'number' && !isNaN(n) ? n : NaN
    }
    
    // Filtrar solo filas relevadas
    const relevadosData = dataToAnalyze.filter(row => {
      if (relevadorIndex === -1) return false
      return String(row[relevadorIndex] || '').trim() !== ''
    })
    
    // Relevados que además tienen lat/long válidos (no vacíos, no 0)
    const relevadosConCoordenadas = latIndex !== -1 && lngIndex !== -1
      ? relevadosData.filter(row => {
          const lat = parseCoord(row[latIndex])
          const lng = parseCoord(row[lngIndex])
          return !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0
        }).length
      : 0
    
    // Campos a analizar (misma lista que statsFields para consistencia)
    const fieldsToAnalyze = [
      'estado kiosco',
      'dias de atención',
      'dias de atencion',
      'horario',
      'escaparate',
      'ubicación',
      'ubicacion',
      'fachada puesto',
      'fachada de puesto',
      'venta productos no editoriales',
      'venta de productos no editoriales',
      'suscripciones',
      'mayor venta',
      'utiliza parada online',
      'utiliza parada online?',
      'reparto',
      'paquete digital',
      'tipo de local',
      'competencia'
    ]
    
    // Colores para gráficos
    const colors = [
      '#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6',
      '#EC4899', '#06B6D4', '#84CC16', '#F97316', '#6366F1',
      '#14B8A6', '#A855F7', '#F43F5E', '#22D3EE', '#FB923C'
    ]
    
    headersLower.forEach((header, idx) => {
      if (fieldsToAnalyze.some(f => header.includes(f))) {
        const counts: { [value: string]: number } = {}
        
        relevadosData.forEach(row => {
          const value = String(row[idx] || '').trim()
          if (value) {
            const normalizedValue = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase()
            counts[normalizedValue] = (counts[normalizedValue] || 0) + 1
          }
        })
        
        const entries = Object.entries(counts)
          .map(([label, count], i) => ({
            label,
            count,
            color: colors[i % colors.length]
          }))
          .sort((a, b) => b.count - a.count)
        
        if (entries.length > 0) {
          results.push({
            fieldName: headersToUse[idx],
            data: entries
          })
        }
      }
    })
    
    const total = dataToAnalyze.length
    const relevados = relevadosData.length
    const pendientes = total - relevados
    
    return { stats: results, total, relevados, pendientes, relevadosConCoordenadas }
  }

  // Descargar estadísticas como CSV/Excel
  const downloadStatsAsExcel = () => {
    let csvContent = '\uFEFF' // BOM para UTF-8 en Excel
    
    // Determinar si usar estadísticas por hoja (admin) o las normales
    const isAdminWithSheetStats = sheetData?.permissions?.isAdmin && Object.keys(allSheetsStats).length > 0
    
    if (isAdminWithSheetStats) {
      // Descargar estadísticas de la hoja seleccionada en el modal
      const { stats, total, relevados, pendientes } = getFieldStatsForSheet(statsSelectedSheet)
      
      if (stats.length === 0) return
      
      const sheetLabel = statsSelectedSheet === 'Total' ? 'Total (todas las hojas)' : statsSelectedSheet
      
      csvContent += 'ESTADÍSTICAS DE PDV RELEVADOS\n'
      csvContent += `Hoja:,${sheetLabel}\n`
      csvContent += `Fecha de generación:,${new Date().toLocaleDateString('es-AR')}\n`
      csvContent += `Total PDV:,${total}\n`
      csvContent += `PDV Relevados:,${relevados}\n`
      csvContent += `PDV Pendientes:,${pendientes}\n`
      csvContent += `Progreso:,${total > 0 ? Math.round((relevados / total) * 100) : 0}%\n`
      csvContent += '\n'
      
      stats.forEach(fieldStat => {
        const fieldTotal = fieldStat.data.reduce((sum, d) => sum + d.count, 0)
        
        csvContent += `${fieldStat.fieldName}\n`
        csvContent += 'Opción,Cantidad,Porcentaje\n'
        
        fieldStat.data.forEach(d => {
          const percent = Math.round((d.count / fieldTotal) * 100)
          const label = d.label.includes(',') ? `"${d.label}"` : d.label
          csvContent += `${label},${d.count},${percent}%\n`
        })
        
        csvContent += `Total,${fieldTotal},100%\n`
        csvContent += '\n'
      })
      
      const safeSheetName = statsSelectedSheet.replace(/[^a-zA-Z0-9]/g, '_')
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
      const link = document.createElement('a')
      const url = URL.createObjectURL(blob)
      link.setAttribute('href', url)
      link.setAttribute('download', `Estadisticas_${safeSheetName}_${new Date().toISOString().split('T')[0]}.csv`)
      link.style.visibility = 'hidden'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    } else {
      // Estadísticas normales (no admin)
      const stats = getFieldStats()
      if (stats.length === 0) return
      
      const { relevadorIndex } = getAutoFillIndexes()
      const totalRelevados = sheetData?.data.filter(row => {
        if (relevadorIndex === -1) return false
        return String(row[relevadorIndex] || '').trim() !== ''
      }).length || 0
      
      csvContent += 'ESTADÍSTICAS DE PDV RELEVADOS\n'
      csvContent += `Fecha de generación:,${new Date().toLocaleDateString('es-AR')}\n`
      csvContent += `Total PDV:,${sheetData?.data.length || 0}\n`
      csvContent += `PDV Relevados:,${totalRelevados}\n`
      csvContent += `PDV Pendientes:,${(sheetData?.data.length || 0) - totalRelevados}\n`
      csvContent += '\n'
      
      stats.forEach(fieldStat => {
        const total = fieldStat.data.reduce((sum, d) => sum + d.count, 0)
        
        csvContent += `${fieldStat.fieldName}\n`
        csvContent += 'Opción,Cantidad,Porcentaje\n'
        
        fieldStat.data.forEach(d => {
          const percent = Math.round((d.count / total) * 100)
          const label = d.label.includes(',') ? `"${d.label}"` : d.label
          csvContent += `${label},${d.count},${percent}%\n`
        })
        
        csvContent += `Total,${total},100%\n`
        csvContent += '\n'
      })
      
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
      const link = document.createElement('a')
      const url = URL.createObjectURL(blob)
      link.setAttribute('href', url)
      link.setAttribute('download', `Estadisticas_PDV_${new Date().toISOString().split('T')[0]}.csv`)
      link.style.visibility = 'hidden'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }
  }

  // Descargar reporte de la hoja actual como Excel/CSV
  const downloadSheetReport = () => {
    if (!sheetData || !sheetData.headers || sheetData.data.length === 0) return
    
    let csvContent = '\uFEFF' // BOM para UTF-8 en Excel
    
    // Nombre de la hoja para el archivo
    const sheetName = adminSelectedSheet || 'Hoja_1'
    const safeSheetName = sheetName.replace(/[^a-zA-Z0-9]/g, '_')
    
    // Agregar headers
    const escapedHeaders = sheetData.headers.map(h => {
      const value = String(h || '')
      return value.includes(',') || value.includes('"') || value.includes('\n') 
        ? `"${value.replace(/"/g, '""')}"` 
        : value
    })
    csvContent += escapedHeaders.join(',') + '\n'
    
    // Agregar datos
    sheetData.data.forEach(row => {
      const escapedRow = row.map((cell: any) => {
        const value = String(cell || '')
        return value.includes(',') || value.includes('"') || value.includes('\n') 
          ? `"${value.replace(/"/g, '""')}"` 
          : value
      })
      csvContent += escapedRow.join(',') + '\n'
    })
    
    // Crear y descargar el archivo
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
    const link = document.createElement('a')
    const url = URL.createObjectURL(blob)
    link.setAttribute('href', url)
    link.setAttribute('download', `Reporte_${safeSheetName}_${new Date().toISOString().split('T')[0]}.csv`)
    link.style.visibility = 'hidden'
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  // Descargar cuestionario en blanco para PDV nuevos
  const downloadCuestionario = () => {
    const today = new Date().toLocaleDateString('es-AR', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    })
    
    const cuestionarioHTML = `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Cuestionario de Relevamiento PDV - Clarín</title>
  <style>
    @page {
      size: A4;
      margin: 15mm;
    }
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }
    body {
      font-family: Arial, sans-serif;
      font-size: 11px;
      line-height: 1.4;
      color: #333;
      padding: 10px;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 3px solid #E31837;
      padding-bottom: 10px;
      margin-bottom: 15px;
    }
    .header h1 {
      color: #E31837;
      font-size: 18px;
      margin: 0;
    }
    .header .logo {
      font-size: 24px;
      font-weight: bold;
      color: #E31837;
    }
    .header .date {
      font-size: 10px;
      color: #666;
    }
    .section {
      margin-bottom: 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      padding: 10px;
    }
    .section-title {
      background: #E31837;
      color: white;
      padding: 5px 10px;
      margin: -10px -10px 10px -10px;
      border-radius: 5px 5px 0 0;
      font-size: 12px;
      font-weight: bold;
    }
    .field-row {
      display: flex;
      margin-bottom: 8px;
      gap: 10px;
    }
    .field {
      flex: 1;
    }
    .field-full {
      width: 100%;
      margin-bottom: 8px;
    }
    .field label {
      display: block;
      font-weight: bold;
      font-size: 10px;
      color: #555;
      margin-bottom: 3px;
    }
    .field input, .field-full input {
      width: 100%;
      border: 1px solid #ccc;
      border-radius: 4px;
      padding: 6px 8px;
      font-size: 11px;
      background: #fafafa;
    }
    .field-line {
      border-bottom: 1px solid #333;
      min-height: 22px;
      margin-top: 3px;
    }
    .checkbox-group {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 5px;
    }
    .checkbox-item {
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: 10px;
    }
    .checkbox-item input[type="checkbox"] {
      width: 14px;
      height: 14px;
    }
    .radio-group {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 5px;
    }
    .radio-item {
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: 10px;
    }
    .radio-item input[type="radio"] {
      width: 14px;
      height: 14px;
    }
    .notes-box {
      border: 1px solid #ccc;
      border-radius: 4px;
      min-height: 50px;
      padding: 8px;
      background: #fafafa;
      margin-top: 5px;
    }
    .footer {
      margin-top: 15px;
      padding-top: 10px;
      border-top: 1px solid #ddd;
      font-size: 9px;
      color: #666;
      text-align: center;
    }
    .required {
      color: #E31837;
    }
    .two-cols {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }
    .three-cols {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr;
      gap: 10px;
    }
    @media print {
      body { padding: 0; }
      .section { break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <div class="logo">Clarín</div>
      <h1>Cuestionario de Relevamiento PDV</h1>
    </div>
    <div class="date">Fecha de impresión: ${today}</div>
  </div>

  <div class="section">
    <div class="section-title">📍 DATOS DE UBICACIÓN</div>
    <div class="two-cols">
      <div class="field">
        <label>ID del PDV (si existe):</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>Paquete: <span class="required">*</span></label>
        <div class="field-line"></div>
      </div>
    </div>
    <div class="field-full">
      <label>Domicilio completo: <span class="required">*</span></label>
      <div class="field-line"></div>
    </div>
    <div class="three-cols">
      <div class="field">
        <label>Provincia:</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>Partido:</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>Localidad / Barrio:</label>
        <div class="field-line"></div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">🏪 ESTADO DEL PUESTO</div>
    <div class="field-full">
      <label>Estado del Kiosco:</label>
      <div class="radio-group">
        <div class="radio-item"><input type="radio" name="estado"> Abierto</div>
        <div class="radio-item"><input type="radio" name="estado"> Cerrado ahora</div>
        <div class="radio-item"><input type="radio" name="estado"> Abre ocasionalmente</div>
        <div class="radio-item"><input type="radio" name="estado"> Cerrado definitivamente</div>
        <div class="radio-item"><input type="radio" name="estado"> No se encuentra el puesto</div>
        <div class="radio-item"><input type="radio" name="estado"> Zona Peligrosa</div>
      </div>
    </div>
    <div class="two-cols">
      <div class="field">
        <label>Días de atención:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="dias"> Todos los días</div>
          <div class="radio-item"><input type="radio" name="dias"> De L a V</div>
          <div class="radio-item"><input type="radio" name="dias"> Sábado y Domingo</div>
          <div class="radio-item"><input type="radio" name="dias"> 3 veces por semana</div>
          <div class="radio-item"><input type="radio" name="dias"> 4 veces por semana</div>
        </div>
      </div>
      <div class="field">
        <label>Horario:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="horario"> Mañana</div>
          <div class="radio-item"><input type="radio" name="horario"> Tarde</div>
          <div class="radio-item"><input type="radio" name="horario"> Mañana y Tarde</div>
          <div class="radio-item"><input type="radio" name="horario"> Solo reparto/Susc.</div>
        </div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">📊 CARACTERÍSTICAS DEL PUESTO</div>
    <div class="three-cols">
      <div class="field">
        <label>Escaparate:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="escaparate"> Chico</div>
          <div class="radio-item"><input type="radio" name="escaparate"> Mediano</div>
          <div class="radio-item"><input type="radio" name="escaparate"> Grande</div>
        </div>
      </div>
      <div class="field">
        <label>Ubicación:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="ubicacion"> Avenida</div>
          <div class="radio-item"><input type="radio" name="ubicacion"> Barrio</div>
          <div class="radio-item"><input type="radio" name="ubicacion"> Estación Subte/Tren</div>
        </div>
      </div>
      <div class="field">
        <label>Fachada del puesto:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="fachada"> Malo</div>
          <div class="radio-item"><input type="radio" name="fachada"> Regular</div>
          <div class="radio-item"><input type="radio" name="fachada"> Bueno</div>
        </div>
      </div>
    </div>
    <div class="three-cols">
      <div class="field">
        <label>Venta prod. no editoriales: <span class="required">*</span></label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="venta"> Nada</div>
          <div class="radio-item"><input type="radio" name="venta"> Poco</div>
          <div class="radio-item"><input type="radio" name="venta"> Mucho</div>
        </div>
      </div>
      <div class="field">
        <label>Reparto:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="reparto"> Sí</div>
          <div class="radio-item"><input type="radio" name="reparto"> No</div>
          <div class="radio-item"><input type="radio" name="reparto"> Ocasionalmente</div>
        </div>
      </div>
      <div class="field">
        <label>Suscripciones:</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="suscripciones"> Sí</div>
          <div class="radio-item"><input type="radio" name="suscripciones"> No</div>
        </div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">📱 DATOS DE CONTACTO Y DISTRIBUCIÓN</div>
    <div class="two-cols">
      <div class="field">
        <label>Teléfono: <span class="required">*</span> (poner 0 si no se obtiene)</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>N° Vendedor:</label>
        <div class="field-line"></div>
      </div>
    </div>
    <div class="two-cols">
      <div class="field">
        <label>Distribuidora:</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>¿Utiliza Parada Online?</label>
        <div class="radio-group">
          <div class="radio-item"><input type="radio" name="parada"> Sí</div>
          <div class="radio-item"><input type="radio" name="parada"> No</div>
        </div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">📝 OBSERVACIONES Y SUGERENCIAS</div>
    <div class="field-full">
      <label>Sugerencias / Comentarios del PDV:</label>
      <div class="notes-box"></div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">✅ DATOS DEL RELEVAMIENTO</div>
    <div class="three-cols">
      <div class="field">
        <label>Fecha de relevamiento:</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>Relevado por (email):</label>
        <div class="field-line"></div>
      </div>
      <div class="field">
        <label>Firma:</label>
        <div class="field-line"></div>
      </div>
    </div>
  </div>

  <div class="footer">
    <p>Clarín - Relevamiento de Puntos de Venta | Este formulario es para PDV nuevos no listados en el sistema</p>
    <p>Una vez completado, ingrese los datos en el sistema o entregue a su supervisor</p>
  </div>

  <script>
    window.onload = function() {
      window.print();
    }
  </script>
</body>
</html>
    `.trim()

    // Abrir en nueva ventana para imprimir
    const printWindow = window.open('', '_blank')
    if (printWindow) {
      printWindow.document.write(cuestionarioHTML)
      printWindow.document.close()
    }
  }

  // Funciones para Nuevo PDV
  const resetNuevoPdvForm = () => {
    setNuevoPdvData({
      estadoKiosco: 'Abierto',
      paquete: '',
      domicilio: '',
      provincia: 'Buenos Aires',
      partido: '',
      localidad: '',
      nVendedor: '',
      distribuidora: '',
      diasAtencion: '',
      horario: '',
      escaparate: '',
      ubicacion: '',
      fachada: '',
      ventaNoEditorial: '',
      reparto: '',
      suscripciones: '',
      nombreApellido: '',
      mayorVenta: '',
      paradaOnline: '',
      telefono: '',
      correoElectronico: '',
      observaciones: '',
      comentarios: '',
      latitud: '',
      longitud: ''
    })
    setNuevoPdvImagePreview(null)
    setNuevoPdvImageUrl(null)
    setNuevoPdvFieldErrors({})
  }

  // Opciones para el formulario de nuevo PDV (idénticas al de edición)
  const nuevoPdvOptions = {
    estadoKiosco: ['Abierto', 'Ahora es Cafeteria', 'Abierto pero otro rubro', 'Cerrado ahora', 'Abre ocasionalmente', 'Cerrado pero hace reparto', 'Cerrado definitivamente', 'Zona Peligrosa', 'No se encuentra el puesto'],
    diasAtencion: ['Todos los dias', 'De L a V', 'Sabado y Domingo', '3 veces por semana', '4 veces por Semana'],
    horario: ['Mañana', 'Mañana y Tarde', 'Tarde', 'Solo reparto/Susc.'],
    escaparate: ['Chico', 'Mediano', 'Grande'],
    ubicacion: ['Avenida', 'Barrio', 'Estación Subte/Tren'],
    fachada: ['Malo', 'Regular', 'Bueno'],
    ventaNoEditorial: ['Nada', 'Poco', 'Mucho'],
    mayorVenta: ['Mostrador', 'Reparto', 'Suscripciones', 'No sabe / No comparte'],
    reparto: ['Si', 'No', 'Ocasionalmente'],
    suscripciones: ['Si', 'No'],
    paradaOnline: ['Si', 'No', 'No sabe'],
    distribuidora: ['Barracas', 'Belgrano', 'Barrio Norte', 'Zunni', 'Recova', 'Boulogne', 'Del Parque', 'Roca/La Boca', 'Lavalle', 'Mariano Acosta', 'Nueva Era', 'San Isidro', 'Ex Rubbo', 'Ex Lugano', 'Ex Jose C Paz'],
    partido: ['Almirante Brown', 'Avellaneda', 'Berazategui', 'CABA', 'Escobar', 'Esteban Echeverría', 'Ezeiza', 'Florencio Varela', 'Hurlingham', 'Ituzaingó', 'Jose C Paz', 'La Matanza', 'Lanús', 'Lomas de Zamora', 'Malvinas Argentinas', 'Merlo', 'Moreno', 'Morón', 'Pilar', 'Presidente Perón', 'Quilmes', 'San Fernando', 'San Isidro', 'San Martín', 'San Miguel', 'San Vicente', 'Tigre', 'Tres de Febrero', 'Vicente López'],
    localidadesPorPartido: {
      'CABA': ['Agronomía', 'Almagro', 'Balvanera', 'Barracas', 'Belgrano', 'Boedo', 'Caballito', 'Chacarita', 'Coghlan', 'Colegiales', 'Constitución', 'Flores', 'Floresta', 'La Boca', 'La Paternal', 'Liniers', 'Mataderos', 'Monte Castro', 'Montserrat', 'Nueva Pompeya', 'Núñez', 'Palermo', 'Parque Avellaneda', 'Parque Chacabuco', 'Parque Chas', 'Parque Patricios', 'Puerto Madero', 'Recoleta', 'Retiro', 'Saavedra', 'San Cristóbal', 'San Nicolás', 'San Telmo', 'Vélez Sársfield', 'Versalles', 'Villa Crespo', 'Villa del Parque', 'Villa Devoto', 'Villa Gral. Mitre', 'Villa Lugano', 'Villa Luro', 'Villa Ortúzar', 'Villa Pueyrredón', 'Villa Real', 'Villa Riachuelo', 'Villa Santa Rita', 'Villa Soldati', 'Villa Urquiza'],
      'Almirante Brown': ['Adrogué', 'Burzaco', 'Claypole', 'Don Orione', 'Glew', 'José Mármol', 'Longchamps', 'Malvinas Argentinas', 'Ministro Rivadavia', 'Rafael Calzada', 'San Francisco Solano', 'San José'],
      'Avellaneda': ['Avellaneda Centro', 'Crucecita', 'Dock Sud', 'Gerli', 'Piñeyro', 'Sarandí', 'Villa Domínico', 'Wilde'],
      'Berazategui': ['Berazategui Centro', 'El Pato', 'Guillermo Hudson', 'Gutiérrez', 'Juan María Gutiérrez', 'Pereyra', 'Plátanos', 'Ranelagh', 'Sourigues', 'Villa España', 'Villa Mitre'],
      'Escobar': ['Belén de Escobar', 'Garín', 'Ingeniero Maschwitz', 'Loma Verde', 'Maquinista Savio'],
      'Esteban Echeverría': ['9 de Abril', 'Canning', 'El Jagüel', 'Luis Guillón', 'Monte Grande', 'Tristán Suárez'],
      'Ezeiza': ['Aeropuerto Ezeiza', 'Carlos Spegazzini', 'Ezeiza Centro', 'La Unión', 'Tristán Suárez'],
      'Florencio Varela': ['Bosques', 'Don Orione', 'Florencio Varela Centro', 'Gobernador Costa', 'Ingeniero Allan', 'La Capilla', 'San Juan Bautista', 'Santa Rosa', 'Villa Brown', 'Villa San Luis', 'Villa Vatteone', 'Zeballos'],
      'Hurlingham': ['Hurlingham', 'Villa Santos Tesei', 'William Morris'],
      'Ituzaingó': ['Ituzaingó Centro', 'Ituzaingó Sur', 'Villa Udaondo'],
      'Jose C Paz': ['José C. Paz Centro', 'Del Viso', 'Tortuguitas'],
      'La Matanza': ['20 de Junio', 'Aldo Bonzi', 'Ciudad Celina', 'Ciudad Evita', 'González Catán', 'Gregorio de Laferrere', 'Isidro Casanova', 'La Tablada', 'Lomas del Mirador', 'Ramos Mejía', 'San Justo', 'Tablada', 'Tapiales', 'Villa Luzuriaga', 'Villa Madero', 'Virrey del Pino', 'Rafael Castillo'],
      'Lanús': ['Gerli', 'Lanús Este', 'Lanús Oeste', 'Monte Chingolo', 'Remedios de Escalada', 'Valentín Alsina'],
      'Lomas de Zamora': ['Banfield', 'Ingeniero Budge', 'Llavallol', 'Lomas de Zamora Centro', 'Temperley', 'Turdera', 'Villa Fiorito', 'Villa Centenario'],
      'Malvinas Argentinas': ['Adolfo Sourdeaux', 'Area de Promoción El Triángulo', 'Grand Bourg', 'Ingeniero Pablo Nogués', 'Los Polvorines', 'Pablo Nogués', 'Tierras Altas', 'Tortuguitas', 'Villa de Mayo'],
      'Merlo': ['Libertad', 'Mariano Acosta', 'Merlo Centro', 'Parque San Martín', 'Pontevedra', 'San Antonio de Padua'],
      'Moreno': ['Cuartel V', 'Francisco Álvarez', 'La Reja', 'Moreno Centro', 'Paso del Rey', 'Trujui'],
      'Morón': ['Castelar', 'El Palomar', 'Haedo', 'Morón Centro', 'Villa Sarmiento'],
      'Pilar': ['Del Viso', 'Fátima', 'La Lonja', 'Manuel Alberti', 'Manzanares', 'Pilar Centro', 'President Derqui', 'Villa Astolfi', 'Villa Rosa'],
      'Presidente Perón': ['Guernica', 'San Martín'],
      'Quilmes': ['Bernal', 'Bernal Oeste', 'Don Bosco', 'Ezpeleta', 'Quilmes Centro', 'Quilmes Oeste', 'San Francisco Solano', 'Villa La Florida'],
      'San Fernando': ['San Fernando Centro', 'Victoria', 'Virreyes'],
      'San Isidro': ['Acassuso', 'Beccar', 'Boulogne Sur Mer', 'La Horqueta', 'Martínez', 'San Isidro Centro', 'Villa Adelina'],
      'San Martín': ['Billinghurst', 'José León Suárez', 'San Andrés', 'San Martín Centro', 'Villa Ballester', 'Villa Lynch', 'Villa Maipú', 'Villa Zagala'],
      'San Miguel': ['Bella Vista', 'Campo de Mayo', 'Muñiz', 'San Miguel Centro', 'Santa María'],
      'San Vicente': ['Alejandro Korn', 'Domselaar', 'San Vicente Centro'],
      'Tigre': ['Benavídez', 'Don Torcuato', 'El Talar', 'General Pacheco', 'Nordelta', 'Ricardo Rojas', 'Rincón de Milberg', 'Tigre Centro', 'Troncos del Talar'],
      'Tres de Febrero': ['Caseros', 'Ciudadela', 'Ciudad Jardín Lomas del Palomar', 'El Libertador', 'José Ingenieros', 'Loma Hermosa', 'Martín Coronado', 'Once de Septiembre', 'Pablo Podestá', 'Sáenz Peña', 'Santos Lugares', 'Villa Bosch', 'Villa Raffo'],
      'Vicente López': ['Carapachay', 'Florida', 'Florida Oeste', 'La Lucila', 'Munro', 'Olivos', 'Vicente López Centro', 'Villa Adelina', 'Villa Martelli'],
    }
  }

  // Estado para sugerencias de localidad
  const [localidadSugerencias, setLocalidadSugerencias] = useState<string[]>([])
  const [showLocalidadSugerencias, setShowLocalidadSugerencias] = useState(false)

  // Filtrar localidades según lo que escribe el usuario
  const handleLocalidadChange = (value: string) => {
    setNuevoPdvData({...nuevoPdvData, localidad: value})
    
    if (nuevoPdvData.partido && value.length > 0) {
      const localidades = (nuevoPdvOptions.localidadesPorPartido as Record<string, string[]>)[nuevoPdvData.partido] || []
      const filtered = localidades.filter(loc => 
        loc.toLowerCase().includes(value.toLowerCase())
      )
      setLocalidadSugerencias(filtered)
      setShowLocalidadSugerencias(filtered.length > 0)
    } else {
      setLocalidadSugerencias([])
      setShowLocalidadSugerencias(false)
    }
  }

  const selectLocalidad = (localidad: string) => {
    setNuevoPdvData({...nuevoPdvData, localidad})
    setShowLocalidadSugerencias(false)
  }

  const handleNuevoPdvImageUpload = async (e: React.ChangeEvent<HTMLInputElement>, captureLocation: boolean = false) => {
    const file = e.target.files?.[0]
    if (!file || !accessToken) return

    // Limpiar el input file inmediatamente para evitar re-disparos
    const inputElement = e.target

    if (!file.type.startsWith('image/')) {
      alert('Por favor selecciona un archivo de imagen válido')
      inputElement.value = ''
      return
    }

    if (file.size > 32 * 1024 * 1024) {
      alert('La imagen es demasiado grande. Máximo 32MB')
      inputElement.value = ''
      return
    }

    const reader = new FileReader()
    reader.onload = (event) => {
      setNuevoPdvImagePreview(event.target?.result as string)
    }
    reader.readAsDataURL(file)
    
    // Limpiar el input para evitar problemas de re-render
    inputElement.value = ''

    setUploadingNuevoPdvImage(true)
    try {
      // Si se solicita capturar ubicación (cámara), obtenerla
      let location: {latitude: number, longitude: number} | null = null
      if (captureLocation) {
        location = await getCurrentLocation()
        if (location) {
          setNuevoPdvData(prev => ({
            ...prev,
            latitud: location!.latitude.toFixed(6),
            longitud: location!.longitude.toFixed(6)
          }))
        }
      }

      const formData = new FormData()
      formData.append('image', file)

      const response = await fetch('/api/upload', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        body: formData
      })

      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          return
        }
        let errorMessage = 'Error al subir imagen'
        try {
          const errorData = await response.json()
          errorMessage = errorData.error || errorMessage
        } catch {
          if (response.status === 413) {
            errorMessage = 'La imagen es demasiado grande. Máximo 32MB.'
          } else if (response.status >= 500) {
            errorMessage = 'Error del servidor. Intenta de nuevo en unos minutos.'
          }
        }
        throw new Error(errorMessage)
      }

      const data = await response.json()
      if (!data.imageUrl) {
        throw new Error('El servidor no devolvió la URL de la imagen. Intenta de nuevo.')
      }
      setNuevoPdvImageUrl(data.imageUrl)

      // Mensaje de éxito con información de ubicación (usar setTimeout para que se muestre después del re-render)
      const successMessage = location 
        ? `✅ Imagen subida correctamente\n📍 Ubicación capturada: ${location.latitude.toFixed(6)}, ${location.longitude.toFixed(6)}`
        : captureLocation 
          ? '✅ Imagen subida correctamente\n⚠️ No se pudo obtener la ubicación GPS'
          : null
      
      if (successMessage) {
        setTimeout(() => toast.success(successMessage.replace(/^✅\s*/, '').replace(/\n/g, ' ')), 100)
      }
    } catch (error: any) {
      console.error('Error uploading image:', error)
      toast.error('Error al subir la imagen: ' + error.message)
      setNuevoPdvImagePreview(null)
    } finally {
      setUploadingNuevoPdvImage(false)
    }
  }

  const handleSaveNuevoPdv = async () => {
    if (!accessToken) return

    // Validaciones
    const errores: string[] = []
    if (!nuevoPdvData.paquete.trim()) errores.push('- Paquete')
    if (!nuevoPdvData.domicilio.trim()) errores.push('- Domicilio')
    if (!nuevoPdvData.ventaNoEditorial.trim()) errores.push('- Venta productos no editoriales')
    if (!nuevoPdvData.telefono.trim()) errores.push('- Teléfono (poner 0 si no se obtiene)')

    if (errores.length > 0) {
      const fieldMap: Record<string, string> = {}
      if (!nuevoPdvData.paquete.trim()) fieldMap.paquete = 'Paquete es obligatorio'
      if (!nuevoPdvData.domicilio.trim()) fieldMap.domicilio = 'Domicilio es obligatorio'
      if (!nuevoPdvData.ventaNoEditorial.trim()) fieldMap.ventaNoEditorial = 'Seleccione una opción'
      if (!nuevoPdvData.telefono.trim()) fieldMap.telefono = 'Teléfono es obligatorio (0 si no se obtiene)'
      setNuevoPdvFieldErrors(fieldMap)
      toast.error('Complete los campos obligatorios marcados')
      return
    }
    setNuevoPdvFieldErrors({})

    const confirmSave = window.confirm('¿Estás seguro de que deseas dar de alta este nuevo PDV?')
    if (!confirmSave) return

    setSavingNuevoPdv(true)
    try {
      // Normalizar Partido y Localidad a Title Case
      const normalizedPdvData = {
        ...nuevoPdvData,
        partido: nuevoPdvData.partido ? toTitleCase(nuevoPdvData.partido.trim()) : '',
        localidad: nuevoPdvData.localidad ? toTitleCase(nuevoPdvData.localidad.trim()) : '',
        imageUrl: nuevoPdvImageUrl || ''
      }
      
      const response = await fetch('/api/alta-pdv', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          pdvData: normalizedPdvData
        })
      })

      if (!response.ok) {
        if (response.status === 401) {
          handleSessionExpired()
          return
        }
        const errorData = await response.json().catch(() => ({}))
        throw new Error(errorData.error || 'Error al guardar el PDV')
      }

      const data = await response.json()
      toast.success(data.message || 'PDV guardado correctamente')
      
      // Cerrar modales y resetear
      setShowNuevoPdvForm(false)
      setShowNuevoPdvModal(false)
      resetNuevoPdvForm()

    } catch (error: any) {
      console.error('Error saving nuevo PDV:', error)
      toast.error('Error al guardar el PDV: ' + error.message)
    } finally {
      setSavingNuevoPdv(false)
    }
  }

  const filteredData = sheetData?.data.filter(row => {
    // Omitir filas sin ID (evita filas en blanco al inicio cuando se usa "Relevamiento: Todos")
    const rowId = String(row?.[0] ?? '').trim()
    if (!rowId) return false

    // Filtro por búsqueda de texto
    if (searchTerm.trim()) {
      const searchValue = searchTerm.trim().toLowerCase()
      
      if (searchType === 'id') {
        // Exact match for ID (first column)
        const rowId = String(row[0] || '').toLowerCase().trim()
        if (rowId !== searchValue) return false
      } else if (searchType === 'paquete') {
        // Flexible match for Paquete column (contains)
        const paqueteIndex = getPaqueteIndex()
        if (paqueteIndex !== -1) {
          const paqueteValue = String(row[paqueteIndex] || '').toLowerCase().trim()
          if (!paqueteValue.includes(searchValue)) return false
        }
      } else if (searchType === 'direccion') {
        // Búsqueda por Dirección/Domicilio (contains)
        const domicilioIndex = getDomicilioIndex()
        if (domicilioIndex !== -1) {
          const direccionValue = String(row[domicilioIndex] || '').toLowerCase().trim()
          if (!direccionValue.includes(searchValue)) return false
        }
      }
    }
    
    // Filtro por estado de relevamiento y por coordenadas (censados mapeados / sin mapear)
    if (filterRelevado !== 'todos') {
      const { relevadorIndex } = getAutoFillIndexes()
      const relevadorValue = relevadorIndex !== -1 ? String(row[relevadorIndex] || '').trim() : ''
      const isRelevado = relevadorValue !== ''

      const { latIndex, lngIndex } = getLatLngIndexes()
      const latStr = (latIndex !== -1 && row[latIndex] != null) ? String(row[latIndex]).trim().replace(',', '.') : ''
      const lngStr = (lngIndex !== -1 && row[lngIndex] != null) ? String(row[lngIndex]).trim().replace(',', '.') : ''
      const lat = latStr ? parseFloat(latStr) : NaN
      const lng = lngStr ? parseFloat(lngStr) : NaN
      const hasCoords = !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0
      
      if (filterRelevado === 'relevados' && !isRelevado) return false
      if (filterRelevado === 'no_relevados' && isRelevado) return false
      if (filterRelevado === 'censados_sin_mapear') {
        if (!isRelevado || hasCoords) return false
      }
      if (filterRelevado === 'censados_mapeados') {
        if (!isRelevado || !hasCoords) return false
      }
    }

    // Filtro por Partido (case-insensitive)
    if (filterPartido) {
      const partidoIdx = getPartidoIndex()
      if (partidoIdx !== -1) {
        const rowPartido = String(row[partidoIdx] || '').trim().toLowerCase()
        if (rowPartido !== filterPartido.toLowerCase()) return false
      }
    }

    // Filtro por Localidad/Barrio (case-insensitive)
    if (filterLocalidad) {
      const localidadIdx = getLocalidadIndex()
      if (localidadIdx !== -1) {
        const rowLocalidad = String(row[localidadIdx] || '').trim().toLowerCase()
        if (rowLocalidad !== filterLocalidad.toLowerCase()) return false
      }
    }
    
    return true
  }) || []

  // Ordenar datos si hay una columna seleccionada
  const sortedData = sortColumn !== null 
    ? [...filteredData].sort((a, b) => {
        const valueA = String(a[sortColumn] || '').trim().toLowerCase()
        const valueB = String(b[sortColumn] || '').trim().toLowerCase()
        
        // Intentar ordenar numéricamente si ambos son números
        const numA = parseFloat(valueA)
        const numB = parseFloat(valueB)
        
        if (!isNaN(numA) && !isNaN(numB)) {
          return sortDirection === 'asc' ? numA - numB : numB - numA
        }
        
        // Ordenar alfabéticamente
        if (sortDirection === 'asc') {
          return valueA.localeCompare(valueB, 'es')
        } else {
          return valueB.localeCompare(valueA, 'es')
        }
      })
    : filteredData

  // Función para manejar clic en encabezado de columna
  const handleColumnSort = (columnIndex: number) => {
    if (sortColumn === columnIndex) {
      // Si ya está ordenando por esta columna, cambiar dirección
      if (sortDirection === 'asc') {
        setSortDirection('desc')
      } else {
        // Si ya estaba en desc, quitar ordenamiento
        setSortColumn(null)
        setSortDirection('asc')
      }
    } else {
      // Nueva columna, ordenar ascendente
      setSortColumn(columnIndex)
      setSortDirection('asc')
    }
  }

  // Pagination logic (usa sortedData que ya incluye filtrado + ordenamiento)
  const totalPages = Math.ceil(sortedData.length / itemsPerPage)
  const startIndex = (currentPage - 1) * itemsPerPage
  const endIndex = startIndex + itemsPerPage
  const paginatedData = sortedData.slice(startIndex, endIndex)
  
  useEffect(() => {
    setCurrentPage(1)
  }, [searchTerm, searchType, filterRelevado, filterPartido, filterLocalidad])

  // Limpiar localidad cuando cambia el partido
  useEffect(() => {
    setFilterLocalidad('')
  }, [filterPartido])

  // Efecto para mostrar un tip aleatorio durante la carga
  useEffect(() => {
    if (loadingData) {
      const randomTip = loadingTips[Math.floor(Math.random() * loadingTips.length)]
      setCurrentTip(randomTip)
    }
  }, [loadingData])

  const getPageNumbers = () => {
    const pages: (number | string)[] = []
    
    if (totalPages <= 7) {
      for (let i = 1; i <= totalPages; i++) pages.push(i)
    } else {
      pages.push(1)
      if (currentPage > 3) pages.push('...')
      const start = Math.max(2, currentPage - 1)
      const end = Math.min(totalPages - 1, currentPage + 1)
      for (let i = start; i <= end; i++) pages.push(i)
      if (currentPage < totalPages - 2) pages.push('...')
      pages.push(totalPages)
    }
    
    return pages
  }

  // No mostrar login hasta haber comprobado si hay sesión guardada (evita flash al volver desde /mapa)
  if (!sessionChecked) {
    return (
      <div className="login-container session-check-screen">
        <div className="session-check-content">
          <div className="spinner" />
          <p>Verificando sesión...</p>
        </div>
      </div>
    )
  }

  // Login Screen
  if (!isAuthorized) {
    return (
      <LoginScreen
        sessionExpired={sessionExpired}
        onDismissSessionExpired={() => setSessionExpired(false)}
        isReady={isReady}
        isLoading={isLoading}
        onAuthClick={handleAuthClick}
      />
    )
  }

  // Dashboard Screen
  return (
    <div className="dashboard">
      {/* Mobile Reload Button */}
      <button 
        className={`mobile-reload-btn ${loadingData ? 'loading' : ''}`}
        onClick={() => {
          if (accessToken && !loadingData) {
            loadSheetData(accessToken, adminSelectedSheet || '')
          }
        }}
        disabled={loadingData}
        title="Recargar datos"
      >
        🔄
      </button>

      {/* Popup de Bienvenida / Mensaje Importante */}
      {showWelcomePopup && (
        <div className="modal-overlay welcome-popup-overlay" onClick={() => handleCloseWelcomePopup(false)}>
          <div className="welcome-popup-content" onClick={e => e.stopPropagation()}>
            <div className="welcome-popup-icon">
              ⚠️
            </div>
            <div className="welcome-popup-header">
              <h2>¡IMPORTANTE!</h2>
            </div>
            <div className="welcome-popup-body">
              <p className="welcome-popup-message">
                En caso de encontrarte con un puesto que ahora es <strong>"Café"</strong> u otro rubro diferente, 
                por favor indicarlo en el campo <strong>Observaciones</strong> del formulario de edición 
                o cargar la información a través de <strong>"+ Nuevo PDV"</strong>.
              </p>
              <div className="welcome-popup-example">
                <span className="example-label">Ejemplo de observación:</span>
                <span className="example-text">"El puesto ahora es una cafetería / verdulería / etc."</span>
              </div>
            </div>
            <div className="welcome-popup-footer">
              <button 
                className="btn-dont-show-again"
                onClick={() => handleCloseWelcomePopup(true)}
              >
                No mostrar de nuevo
              </button>
              <button 
                className="btn-accept-welcome"
                onClick={() => handleCloseWelcomePopup(false)}
              >
                ✓ Entendido
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Campos Obligatorios Faltantes */}
      {showMissingFieldsModal && missingFields.length > 0 && (
        <div className="modal-overlay" onClick={() => setShowMissingFieldsModal(false)}>
          <div className="modal-content modal-missing-fields" onClick={e => e.stopPropagation()}>
            <div className="modal-header modal-header-warning">
              <h2>⚠️ Campos Obligatorios</h2>
              <button className="modal-close" onClick={() => setShowMissingFieldsModal(false)}>×</button>
            </div>
            <div className="modal-body">
              <p className="missing-fields-intro">
                Por favor complete los siguientes campos para poder guardar:
              </p>
              <div className="missing-fields-form">
                {missingFields.map((field, idx) => (
                  <div key={idx} className="missing-field-item">
                    <label className="missing-field-label">
                      {field.name}
                      {field.hint && <span className="missing-field-hint">({field.hint})</span>}
                    </label>
                    {pendingSaveAction === 'edit' && field.name === 'Venta productos no editoriales' ? (
                      <select
                        className="missing-field-input"
                        value={editedValues[field.index] || ''}
                        onChange={(e) => {
                          const newValues = [...editedValues]
                          newValues[field.index] = e.target.value
                          setEditedValues(newValues)
                        }}
                        autoFocus={idx === 0}
                      >
                        <option value="">-- Seleccionar opción --</option>
                        <option value="Nada">Nada</option>
                        <option value="Poco">Poco</option>
                        <option value="Mucho">Mucho</option>
                        <option value="Puesto Cerrado">Puesto Cerrado</option>
                        <option value="Cerrado Definitivamente">Cerrado Definitivamente</option>
                        <option value="Zona Peligrosa">Zona Peligrosa</option>
                        <option value="No se encuentra el puesto">No se encuentra el puesto</option>
                      </select>
                    ) : (
                      <input
                        type="text"
                        className="missing-field-input"
                        value={pendingSaveAction === 'edit' ? (editedValues[field.index] || '') : ''}
                        onChange={(e) => {
                          if (pendingSaveAction === 'edit') {
                            const newValues = [...editedValues]
                            newValues[field.index] = e.target.value
                            setEditedValues(newValues)
                          }
                        }}
                        placeholder={field.hint || `Ingrese ${field.name.toLowerCase()}`}
                        autoFocus={idx === 0}
                      />
                    )}
                  </div>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button 
                className="btn-secondary"
                onClick={() => setShowMissingFieldsModal(false)}
              >
                Cancelar
              </button>
              <button 
                className="btn-primary"
                onClick={() => {
                  // Verificar que todos los campos estén completos
                  const allFilled = missingFields.every(field => {
                    if (pendingSaveAction === 'edit') {
                      return String(editedValues[field.index] || '').trim() !== ''
                    }
                    return true
                  })
                  
                  if (!allFilled) {
                    alert('Por favor complete todos los campos')
                    return
                  }
                  
                  setShowMissingFieldsModal(false)
                  if (pendingSaveAction === 'edit') {
                    handleSaveRow(true) // Skip validation since we just filled the fields
                  }
                }}
              >
                ✓ Completar y Guardar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Edit Modal */}
      {editingRow !== null && sheetData && (() => {
        const { fechaIndex, relevadorIndex } = getAutoFillIndexes()
        const today = new Date().toLocaleDateString('es-AR', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })
        
        return (
          <div className="modal-overlay" onClick={handleCancelEdit}>
            <div className="modal-content" onClick={e => e.stopPropagation()}>
              <div className="modal-header">
                <h2>Editar Registro</h2>
                <button className="modal-close" onClick={handleCancelEdit}>×</button>
              </div>
              <div className="modal-body">
                {/* Selector de Puesto Activo/Cerrado */}
                <div className="puesto-status-selector">
                  <label>¿Cuál es el estado del puesto?</label>
                  <div className="puesto-status-buttons">
                    <button
                      type="button"
                      className={`puesto-btn puesto-btn-abierto ${puestoStatus === 'abierto' ? 'active' : ''}`}
                      onClick={() => handlePuestoStatusChange('abierto')}
                    >
                      ✓ Puesto Activo
                    </button>
                    <button
                      type="button"
                      className={`puesto-btn puesto-btn-cerrado ${puestoStatus === 'cerrado' ? 'active' : ''}`}
                      onClick={() => handlePuestoStatusChange('cerrado')}
                    >
                      ✗ Puesto Cerrado DEFINITIVAMENTE
                    </button>
                    <button
                      type="button"
                      className={`puesto-btn puesto-btn-no-encontrado ${puestoStatus === 'no_encontrado' ? 'active' : ''}`}
                      onClick={() => handlePuestoStatusChange('no_encontrado')}
                    >
                      ? No se encontró el puesto
                    </button>
                    <button
                      type="button"
                      className={`puesto-btn puesto-btn-peligrosa ${puestoStatus === 'zona_peligrosa' ? 'active' : ''}`}
                      onClick={() => handlePuestoStatusChange('zona_peligrosa')}
                    >
                      ⚠ Zona Peligrosa
                    </button>
                    <button
                      type="button"
                      className={`puesto-btn puesto-btn-cafeteria ${puestoStatus === 'cafeteria' ? 'active' : ''}`}
                      onClick={() => handlePuestoStatusChange('cafeteria')}
                    >
                      ☕ Cafeteria
                    </button>
                  </div>
                  {puestoStatus === 'cerrado' && (
                    <div className="puesto-cerrado-notice">
                      <span className="notice-icon">⚠️</span>
                      <span>Solo el campo "Estado Kiosco" se ha cambiado a "Cerrado definitivamente". El resto de los campos mantienen sus valores originales.</span>
                    </div>
                  )}
                  {puestoStatus === 'no_encontrado' && (
                    <div className="puesto-cerrado-notice puesto-no-encontrado-notice">
                      <span className="notice-icon">❓</span>
                      <span>Solo el campo "Estado Kiosco" se ha cambiado a "No se encuentra el puesto". El resto de los campos mantienen sus valores originales.</span>
                    </div>
                  )}
                  {puestoStatus === 'zona_peligrosa' && (
                    <div className="puesto-cerrado-notice puesto-peligrosa-notice">
                      <span className="notice-icon">🚨</span>
                      <span>Solo el campo "Estado Kiosco" se ha cambiado a "Zona Peligrosa". El resto de los campos mantienen sus valores originales.</span>
                    </div>
                  )}
                  {puestoStatus === 'cafeteria' && (
                    <div className="puesto-cerrado-notice puesto-cafeteria-notice">
                      <span className="notice-icon">☕</span>
                      <span>El campo "Estado Kiosco" se ha cambiado a "Ahora es Cafeteria". El resto de los campos mantienen sus valores originales.</span>
                    </div>
                  )}
                </div>

                <div className="auto-fill-notice">
                  <span className="notice-icon">ℹ️</span>
                  <span>Al guardar, se registrará automáticamente la fecha ({today}) y tu email como relevador.</span>
                </div>
                <div className="edit-form">
                  {sheetData.headers.map((header, idx) => {
                    const isIdField = idx === 0
                    const isFechaField = idx === fechaIndex
                    const isRelevadorField = idx === relevadorIndex
                    const isAutoField = isFechaField || isRelevadorField
                    
                    // Detectar si es el campo Estado Kiosco
                    const headerLower = header.toLowerCase().trim()
                    
                    // Ocultar campo IMG (se maneja con el componente de subida de imagen)
                    const isImgField = headerLower === 'img' || headerLower === 'imagen'
                    if (isImgField) return null
                    
                    // Ocultar campos de latitud y longitud (se guardan automáticamente con la foto)
                    const isLatLngField = headerLower === 'latitud' || headerLower === 'lat' || 
                                          headerLower === 'longitud' || headerLower === 'lng' || headerLower === 'long'
                    if (isLatLngField) return null
                    
                    // Ocultar campo comentario de foto
                    const isComentarioFotoField = headerLower.includes('comentario') && headerLower.includes('foto')
                    if (isComentarioFotoField) return null
                    
                    // Ocultar campo DISPOSITIVO (se guarda automáticamente)
                    const isDispositivoField = headerLower === 'dispositivo' || headerLower.includes('dispositivo')
                    if (isDispositivoField) return null
                    
                    const isEstadoKioscoField = headerLower.includes('estado') && headerLower.includes('kiosco')
                    
                    // Detectar si es el campo Días de atención
                    const isDiasAtencionField = headerLower.includes('dias de atención') || headerLower.includes('dias de atencion') || headerLower === 'dias de atención' || headerLower === 'dias de atencion'
                    
                    // Detectar si es el campo Horario
                    const isHorarioField = headerLower === 'horario' || headerLower === 'horario:'
                    
                    // Detectar si es el campo Escaparate
                    const isEscaparateField = headerLower === 'escaparate' || headerLower === 'escaparate:'
                    
                    // Detectar si es el campo Ubicación
                    const isUbicacionField = headerLower === 'ubicacion' || headerLower === 'ubicación' || headerLower === 'ubicacion:' || headerLower === 'ubicación:'
                    
                    // Detectar si es el campo Fachada de puesto
                    const isFachadaField = headerLower.includes('fachada') && headerLower.includes('puesto')
                    
                    // Detectar si es el campo Venta de productos no editoriales
                    const isVentaNoEditorialField = headerLower.includes('venta') && headerLower.includes('no editorial')
                    
                    // Detectar si es el campo Reparto
                    const isRepartoField = headerLower === 'reparto' || headerLower === 'reparto:'
                    
                    // Detectar si es el campo Distribuidora
                    const isDistribuidoraField = headerLower === 'distribuidora' || headerLower === 'distribuidora:'
                    
                    // Detectar si es el campo Sugerencias
                    const isSugerenciasField = headerLower.includes('sugerencia') || headerLower.includes('sigeremcia') || headerLower.includes('observacion') || headerLower.includes('observación') || headerLower.includes('comentario')
                    
                    // Corregir nombre mal escrito de Sugerencias
                    const displayHeader = isSugerenciasField && (headerLower.includes('sigeremcia')) 
                      ? 'Sugerencias del PDV' 
                      : header
                    
                    // Detectar si es el campo Teléfono
                    const isTelefonoField = headerLower === 'teléfono' || headerLower === 'telefono' || headerLower === 'teléfono:' || headerLower === 'telefono:' || headerLower.includes('telefono') || headerLower.includes('teléfono')
                    
                    // Detectar si es el campo Provincia (no editable)
                    const isProvinciaField = headerLower === 'provincia' || headerLower === 'provincia:'
                    
                    // Detectar si es el campo Paquete
                    const isPaqueteField = headerLower.includes('paquete')
                    
                    // Campos obligatorios: Paquete es SIEMPRE obligatorio, los demás solo cuando está abierto
                    const isCampoObligatorioSiempre = isPaqueteField
                    const isCampoObligatorioSoloAbierto = isVentaNoEditorialField || isTelefonoField
                    const isCampoObligatorio = isCampoObligatorioSiempre || (isCampoObligatorioSoloAbierto && (puestoStatus === 'abierto' || puestoStatus === 'cafeteria'))
                    
                    const estadoKioscoOptions = [
                      'Abierto',
                      'Ahora es Cafeteria',
                      'Abierto pero otro rubro',
                      'Cerrado ahora',
                      'Abre ocasionalmente',
                      'Cerrado pero hace reparto',
                      'Cerrado definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const diasAtencionOptions = [
                      'Todos los dias',
                      'De L a V',
                      'Sabado y Domingo',
                      '3 veces por semana',
                      '4 veces por Semana',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const horarioOptions = [
                      'Mañana',
                      'Mañana y Tarde',
                      'Tarde',
                      'Solo reparto/Susc.',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const escaparateOptions = [
                      'Chico',
                      'Mediano',
                      'Grande',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const ubicacionOptions = [
                      'Avenida',
                      'Barrio',
                      'Estación Subte/Tren',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const fachadaOptions = [
                      'Malo',
                      'Regular',
                      'Bueno',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const ventaNoEditorialOptions = [
                      'Nada',
                      'Poco',
                      'Mucho',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const repartoOptions = [
                      'Si',
                      'No',
                      'Ocasionalmente',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    const distribuidoraOptions = [
                      'Barracas',
                      'Belgrano',
                      'Barrio Norte',
                      'Zunni',
                      'Recova',
                      'Boulogne',
                      'Del Parque',
                      'Roca/La Boca',
                      'Lavalle',
                      'Mariano Acosta',
                      'Nueva Era',
                      'San Isidro',
                      'Ex Rubbo',
                      'Ex Lugano',
                      'Ex Jose C Paz'
                    ]
                    
                    // Detectar si es el campo Suscripciones
                    const isSuscripcionesField = headerLower === 'suscripciones' || headerLower === 'suscripciones:' || headerLower === 'suscripcion' || headerLower === 'suscripción'
                    
                    const suscripcionesOptions = [
                      'Si',
                      'No',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    // Detectar si es el campo Utiliza Parada Online
                    const isParadaOnlineField = headerLower.includes('parada') && headerLower.includes('online')
                    
                    const paradaOnlineOptions = [
                      'Si',
                      'No',
                      'No sabe',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    // Detectar si es el campo Mayor venta
                    const isMayorVentaField = headerLower.includes('mayor') && headerLower.includes('venta')
                    
                    const mayorVentaOptions = [
                      'Mostrador',
                      'Reparto',
                      'Suscripciones',
                      'No sabe / No comparte',
                      'Puesto Cerrado',
                      'Cerrado Definitivamente',
                      'Zona Peligrosa',
                      'No se encuentra el puesto'
                    ]
                    
                    // Detectar si es el campo Partido
                    const isPartidoField = headerLower === 'partido' || headerLower === 'partido:'
                    
                    const partidoOptions = [
                      'Almirante Brown',
                      'Avellaneda',
                      'Berazategui',
                      'CABA',
                      'Escobar',
                      'Esteban Echeverría',
                      'Ezeiza',
                      'Florencio Varela',
                      'Hurlingham',
                      'Ituzaingó',
                      'Jose C Paz',
                      'La Matanza',
                      'Lanús',
                      'Lomas de Zamora',
                      'Malvinas Argentinas',
                      'Merlo',
                      'Moreno',
                      'Morón',
                      'Pilar',
                      'Presidente Perón',
                      'Quilmes',
                      'San Fernando',
                      'San Isidro',
                      'San Martín',
                      'San Miguel',
                      'San Vicente',
                      'Tigre',
                      'Tres de Febrero',
                      'Vicente López'
                    ]
                    
                    // Detectar si es el campo Localidad/Barrio
                    const isLocalidadField = headerLower === 'localidad' || headerLower === 'localidad:' || 
                                             headerLower === 'barrio' || headerLower === 'barrio:' ||
                                             headerLower === 'localidad/barrio' || headerLower === 'localidad/barrio:' ||
                                             headerLower === 'localidad / barrio' || headerLower === 'localidad / barrio:'
                    
                    // Mapeo de localidades/barrios por partido
                    const localidadesPorPartido: { [key: string]: string[] } = {
                      'Almirante Brown': [
                        'Adrogué',
                        'Burzaco',
                        'Claypole',
                        'Don Orione',
                        'Glew',
                        'José Marmol',
                        'Longchamps',
                        'Malvinas Argentinas',
                        'Ministro Rivadavia',
                        'Rafael Calzada',
                        'San Francisco Solano'
                      ],
                      'Avellaneda': [
                        'Avellaneda',
                        'Crucecita',
                        'Dock Sud',
                        'Gerli',
                        'Piñeyro',
                        'Sarandí',
                        'Villa Domínico',
                        'Wilde'
                      ],
                      'Berazategui': [
                        'Berazategui',
                        'El Pato',
                        'Guillermo Hudson',
                        'Gutiérrez',
                        'Pereyra',
                        'Plátanos',
                        'Ranelagh',
                        'Sourigues',
                        'Villa España'
                      ],
                      'CABA': [
                        'Agronomía',
                        'Almagro',
                        'Balvanera',
                        'Barracas',
                        'Belgrano',
                        'Boedo',
                        'Caballito',
                        'Chacarita',
                        'Coghlan',
                        'Colegiales',
                        'Constitución',
                        'Flores',
                        'Floresta',
                        'La Boca',
                        'La Paternal',
                        'Liniers',
                        'Mataderos',
                        'Monte Castro',
                        'Montserrat',
                        'Nueva Pompeya',
                        'Núñez',
                        'Palermo',
                        'Parque Avellaneda',
                        'Parque Chacabuco',
                        'Parque Chas',
                        'Parque Patricios',
                        'Puerto Madero',
                        'Recoleta',
                        'Retiro',
                        'Saavedra',
                        'San Cristóbal',
                        'San Nicolás',
                        'San Telmo',
                        'Velez Sarsfield',
                        'Versalles',
                        'Villa Crespo',
                        'Villa Del Parque',
                        'Villa Devoto',
                        'Villa General Mitre',
                        'Villa Lugano',
                        'Villa Luro',
                        'Villa Ortúzar',
                        'Villa Pueyrredon',
                        'Villa Real',
                        'Villa Riachuelo',
                        'Villa Santa Rita',
                        'Villa Soldati',
                        'Villa Urquiza'
                      ],
                      'Escobar': [
                        'Escobar',
                        'Garin',
                        'Ingeniero Maschiwitz',
                        'Maquinista Savio',
                        'Matheu'
                      ],
                      'Esteban Echeverría': [
                        '9 de Abril',
                        'El Jagüel',
                        'Luis Guillón',
                        'Monte Grande'
                      ],
                      'Ezeiza': [
                        'Canning',
                        'Carlos Spegazzini',
                        'Ezeiza',
                        'La Unión',
                        'Tristán Suárez'
                      ],
                      'Florencio Varela': [
                        'Bosques',
                        'Florencio Varela',
                        'Gobernador Costa',
                        'Ingeniero Allan',
                        'La Capilla',
                        'Villa Brown',
                        'Villa San Luis',
                        'Villa Santa Rosa',
                        'Villa Vattone',
                        'Zeballos'
                      ],
                      'Hurlingham': [
                        'Hurlingham',
                        'Villa Tesei',
                        'William Morris'
                      ],
                      'Ituzaingó': [
                        'Ituzaingó',
                        'Parque Leloir',
                        'Villa Udaondo'
                      ],
                      'Jose C Paz': [
                        'Jose C Paz'
                      ],
                      'La Matanza': [
                        '20 de Junio',
                        'Aldo Bonzi',
                        'Ciudad Evita',
                        'Gonzalez Catan',
                        'Isidro Casanova',
                        'La Tablada',
                        'Laferrere',
                        'Lomas del Mirador',
                        'Rafael Castillo',
                        'Ramos Mejía',
                        'San Justo',
                        'Tapiales',
                        'Villa Celina',
                        'Villa Luzuriaga',
                        'Villa Madero',
                        'Virrey del Pino'
                      ],
                      'Lanús': [
                        'Gerli',
                        'Lanús Este',
                        'Lanús Oeste',
                        'Monte Chingolo',
                        'Remedios de Escalada',
                        'Valentín Alsina',
                        'Villa Caraza',
                        'Villa Diamante',
                        'Villa Industriales'
                      ],
                      'Lomas de Zamora': [
                        'Banfield',
                        'Ingeniero Budge',
                        'Llavallol',
                        'Lomas de Zamora',
                        'Temperley',
                        'Turdera',
                        'Villa Centenario',
                        'Villa Fiorito'
                      ],
                      'Malvinas Argentinas': [
                        'Grand Bourg',
                        'Los Polvorines',
                        'Malvinas Argentinas',
                        'Sourdeaux',
                        'Tortuguitas',
                        'Villa De Mayo'
                      ],
                      'Merlo': [
                        'Libertad',
                        'Mariano Acosta',
                        'Merlo',
                        'Parque San Martín',
                        'Pontevedra',
                        'San Antonio de Padua'
                      ],
                      'Moreno': [
                        'Cuartel V',
                        'Francisco Alvarez',
                        'La Reja',
                        'Moreno',
                        'Paso Del Rey',
                        'Trujui'
                      ],
                      'Morón': [
                        'Castelar',
                        'El Palomar',
                        'Haedo',
                        'Morón',
                        'Villa Sarmiento'
                      ],
                      'Pilar': [
                        'Del Viso',
                        'Derqui',
                        'Manuel Alberti',
                        'Pilar',
                        'Villa Rosa'
                      ],
                      'Presidente Perón': [
                        'Guernica',
                        'Numancia'
                      ],
                      'Quilmes': [
                        'Bernal',
                        'Bernal Oeste',
                        'Don Bosco',
                        'Ezpeleta',
                        'Ezpeleta Oeste',
                        'Quilmes',
                        'Quilmes Oeste',
                        'Villa La Florida'
                      ],
                      'San Fernando': [
                        'San Fernando',
                        'Victoria',
                        'Virreyes'
                      ],
                      'San Isidro': [
                        'Acassuso',
                        'Beccar',
                        'Boulogne',
                        'Martínez',
                        'San Isidro',
                        'Villa Adelina'
                      ],
                      'San Martín': [
                        'Billinghurst',
                        'José León Suárez',
                        'San Andrés',
                        'San Martín',
                        'Villa Ballester',
                        'Villa Lynch',
                        'Villa Maipú'
                      ],
                      'San Miguel': [
                        'Bella Vista',
                        'Campo de Mayo',
                        'Muñiz',
                        'San Miguel'
                      ],
                      'San Vicente': [
                        'Alejandro Korn',
                        'Domselaar',
                        'San Vicente'
                      ],
                      'Tigre': [
                        'Benavidez',
                        'Don Torcuato',
                        'El Talar',
                        'General Pacheco',
                        'Ricardo Rojas',
                        'Rincon de Milberg',
                        'Tigre'
                      ],
                      'Tres de Febrero': [
                        'Caseros',
                        'Ciudadela',
                        'El Libertador',
                        'Jose Ingenieros',
                        'Loma Hermosa',
                        'Martín Coronado',
                        'Pablo Podestá',
                        'Saenz Peña',
                        'Santos Lugares',
                        'Villa Bosch',
                        'Villa Raffo'
                      ],
                      'Vicente López': [
                        'Carapachay',
                        'Florida',
                        'Florida Oeste',
                        'La Lucila',
                        'Munro',
                        'Olivos',
                        'Vicente López',
                        'Villa Martelli'
                      ]
                    }
                    
                    // Obtener el índice del campo Partido para leer su valor
                    const partidoIndex = sheetData.headers.findIndex(h => {
                      const hLower = h.toLowerCase().trim()
                      return hLower === 'partido' || hLower === 'partido:'
                    })
                    const selectedPartido = partidoIndex !== -1 ? (editedValues[partidoIndex] || '') : ''
                    
                    // Obtener localidades según el partido seleccionado
                    const localidadOptions = localidadesPorPartido[selectedPartido] || []
                    
                    let displayValue = editedValues[idx] || ''
                    let placeholder = ''
                    
                    if (isFechaField) {
                      displayValue = today
                      placeholder = 'Se completará automáticamente'
                    } else if (isRelevadorField) {
                      displayValue = userEmail || ''
                      placeholder = 'Se completará automáticamente'
                    }
                    
                    // Verificar si este campo se debe auto-rellenar cuando está cerrado/no encontrado/zona peligrosa
                    // Solo bloquear el campo Estado Kiosco en todos los casos
                    const isCampoCerrado = (puestoStatus === 'cerrado' || puestoStatus === 'no_encontrado' || puestoStatus === 'zona_peligrosa') && isEstadoKioscoField
                    
                    return (
                      <div key={idx} className={`edit-field ${isAutoField ? 'auto-field' : ''} ${isCampoCerrado ? 'campo-cerrado' : ''} ${isCampoObligatorio ? 'campo-obligatorio' : ''}`}>
                        <label htmlFor={`edit-field-${idx}`}>
                          {isSugerenciasField ? displayHeader : header}
                          {isAutoField && <span className="auto-badge">Auto</span>}
                          {isCampoObligatorio && <span className="obligatorio-badge">* Obligatorio</span>}
                          {isCampoCerrado && (
                            <span className="auto-badge" style={{
                              background: puestoStatus === 'cerrado' ? '#6B7280' : 
                                         puestoStatus === 'no_encontrado' ? '#F59E0B' : 
                                         puestoStatus === 'zona_peligrosa' ? '#DC2626' : '#DC2626'
                            }}>
                              {puestoStatus === 'cerrado' ? 'Cerrado' : 
                               puestoStatus === 'no_encontrado' ? 'No encontrado' : 
                               puestoStatus === 'zona_peligrosa' ? 'Peligrosa' : 'Auto'}
                            </span>
                          )}
                        </label>
                        {isPartidoField || isLocalidadField ? (() => {
                          const fieldOptions = isPartidoField ? partidoOptions : localidadOptions
                          const currentValue = editedValues[idx] || ''
                          const filterKey = `field_${idx}`
                          const filterText = autocompleteFilter[filterKey] ?? currentValue
                          const isOpen = openAutocomplete === filterKey
                          
                          // Filtrar opciones basadas en el texto ingresado
                          const filteredOptions = fieldOptions.filter(opt => 
                            opt.toLowerCase().includes(filterText.toLowerCase())
                          )
                          
                          // Mensaje especial si es Localidad y no hay partido seleccionado
                          const noPartidoSelected = isLocalidadField && !selectedPartido
                          
                          // Obtener el índice del campo Localidad para limpiarlo cuando cambie el partido
                          const localidadIndex = sheetData.headers.findIndex(h => {
                            const hLower = h.toLowerCase().trim()
                            return hLower === 'localidad' || hLower === 'localidad:' || 
                                   hLower === 'barrio' || hLower === 'barrio:' ||
                                   hLower === 'localidad/barrio' || hLower === 'localidad/barrio:' ||
                                   hLower === 'localidad / barrio' || hLower === 'localidad / barrio:'
                          })
                          
                          return (
                            <div className="autocomplete-container">
                              <input
                                id={`edit-field-${idx}`}
                                type="text"
                                value={filterText}
                                placeholder={
                                  isPartidoField 
                                    ? 'Escribir para buscar partido...' 
                                    : noPartidoSelected 
                                      ? 'Primero seleccione un partido...' 
                                      : 'Escribir para buscar localidad...'
                                }
                                onChange={(e) => {
                                  if (noPartidoSelected) return // No permitir escribir si no hay partido
                                  const newFilter = { ...autocompleteFilter, [filterKey]: e.target.value }
                                  setAutocompleteFilter(newFilter)
                                  setOpenAutocomplete(filterKey)
                                }}
                                onFocus={() => {
                                  if (!noPartidoSelected) {
                                    setOpenAutocomplete(filterKey)
                                  }
                                }}
                                onBlur={() => {
                                  // Delay para permitir click en opciones
                                  setTimeout(() => {
                                    if (openAutocomplete === filterKey) {
                                      setOpenAutocomplete(null)
                                      // Si el texto no coincide exactamente con una opción, mantener el valor anterior
                                      if (!fieldOptions.includes(filterText) && filterText !== currentValue) {
                                        const newFilter = { ...autocompleteFilter, [filterKey]: currentValue }
                                        setAutocompleteFilter(newFilter)
                                      }
                                    }
                                  }, 200)
                                }}
                                className={`autocomplete-input ${isCampoObligatorio ? 'input-obligatorio' : ''} ${noPartidoSelected ? 'autocomplete-disabled' : ''}`}
                                disabled={isCampoCerrado || noPartidoSelected}
                                aria-required={isCampoObligatorio}
                              />
                              {isOpen && filteredOptions.length > 0 && (
                                <div className="autocomplete-dropdown">
                                  {filteredOptions.map((option, optIdx) => (
                                    <div
                                      key={optIdx}
                                      className={`autocomplete-option ${option === currentValue ? 'selected' : ''}`}
                                      onMouseDown={(e) => {
                                        e.preventDefault()
                                        const newValues = [...editedValues]
                                        newValues[idx] = option
                                        setEditedValues(newValues)
                                        const newFilter = { ...autocompleteFilter, [filterKey]: option }
                                        setAutocompleteFilter(newFilter)
                                        setOpenAutocomplete(null)
                                        
                                        // Si se selecciona un partido, limpiar la localidad si no es válida
                                        if (isPartidoField && localidadIndex !== -1) {
                                          const currentLocalidad = newValues[localidadIndex] || ''
                                          const newLocalidades = localidadesPorPartido[option] || []
                                          if (currentLocalidad && !newLocalidades.includes(currentLocalidad)) {
                                            newValues[localidadIndex] = ''
                                            setEditedValues([...newValues])
                                            // Limpiar también el filtro de localidad
                                            const localidadFilterKey = `field_${localidadIndex}`
                                            const updatedFilter = { ...autocompleteFilter, [filterKey]: option, [localidadFilterKey]: '' }
                                            setAutocompleteFilter(updatedFilter)
                                          }
                                        }
                                      }}
                                    >
                                      {option}
                                    </div>
                                  ))}
                                </div>
                              )}
                              {isOpen && filteredOptions.length === 0 && filterText && !noPartidoSelected && (
                                <div className="autocomplete-dropdown">
                                  <div className="autocomplete-no-results">
                                    No se encontraron resultados
                                  </div>
                                </div>
                              )}
                              {noPartidoSelected && (
                                <div className="autocomplete-hint">
                                  ⚠️ Seleccione primero un partido para ver las localidades disponibles
                                </div>
                              )}
                            </div>
                          )
                        })() : isEstadoKioscoField || isDiasAtencionField || isHorarioField || isEscaparateField || isUbicacionField || isFachadaField || isVentaNoEditorialField || isRepartoField || isSuscripcionesField || isParadaOnlineField || isMayorVentaField || isDistribuidoraField ? (() => {
                          const currentOptions = isDiasAtencionField ? diasAtencionOptions : isHorarioField ? horarioOptions : isEscaparateField ? escaparateOptions : isUbicacionField ? ubicacionOptions : isFachadaField ? fachadaOptions : isVentaNoEditorialField ? ventaNoEditorialOptions : isRepartoField ? repartoOptions : isSuscripcionesField ? suscripcionesOptions : isParadaOnlineField ? paradaOnlineOptions : isMayorVentaField ? mayorVentaOptions : isDistribuidoraField ? distribuidoraOptions : estadoKioscoOptions
                          const currentValue = editedValues[idx] || ''
                          const valueExistsInOptions = currentOptions.includes(currentValue) || currentValue === ''
                          
                          return (
                            <select
                              id={`edit-field-${idx}`}
                              value={currentValue}
                              onChange={(e) => {
                                if (!isCampoCerrado) {
                                  const newValues = [...editedValues]
                                  newValues[idx] = e.target.value
                                  // Si elige "Cerrado pero hace reparto" en Estado Kiosco, poner Reparto en "Si"
                                  if (isEstadoKioscoField && e.target.value === 'Cerrado pero hace reparto') {
                                    const repartoIdx = sheetData.headers.findIndex((h: string) => {
                                      const l = String(h || '').toLowerCase().trim()
                                      return l === 'reparto' || l === 'reparto:'
                                    })
                                    if (repartoIdx !== -1) newValues[repartoIdx] = 'Si'
                                  }
                                  setEditedValues(newValues)
                                }
                              }}
                              className="estado-kiosco-select"
                              disabled={isCampoCerrado}
                              aria-required={isCampoObligatorio}
                            >
                              <option value="">-- Seleccionar {isDiasAtencionField ? 'días' : isHorarioField ? 'horario' : isEscaparateField ? 'escaparate' : isUbicacionField ? 'ubicación' : isFachadaField ? 'fachada' : isVentaNoEditorialField ? 'opción' : isRepartoField ? 'reparto' : isSuscripcionesField ? 'opción' : isParadaOnlineField ? 'opción' : isMayorVentaField ? 'opción' : isDistribuidoraField ? 'distribuidora' : 'estado'} --</option>
                              {/* Si el valor actual no está en las opciones, mostrarlo primero */}
                              {!valueExistsInOptions && currentValue && (
                                <option key={currentValue} value={currentValue}>{currentValue}</option>
                              )}
                              {currentOptions.map((option) => (
                                <option key={option} value={option}>{option}</option>
                              ))}
                            </select>
                          )
                        })() : isSugerenciasField ? (
                          <textarea
                            id={`edit-field-${idx}`}
                            value={editedValues[idx] || ''}
                            placeholder=""
                            onChange={(e) => {
                              if (!isCampoCerrado) {
                                const newValues = [...editedValues]
                                newValues[idx] = e.target.value
                                setEditedValues(newValues)
                              }
                            }}
                            disabled={isCampoCerrado}
                            className="sugerencias-textarea"
                            rows={3}
                            aria-required={isCampoObligatorio}
                          />
                        ) : (
                          <input
                            id={`edit-field-${idx}`}
                            type="text"
                            value={isAutoField ? displayValue : (editedValues[idx] || '')}
                            placeholder={isTelefonoField ? 'Ingrese teléfono (poner 0 si no se obtiene)' : placeholder}
                            onChange={(e) => {
                              if (!isIdField && !isAutoField && !isCampoCerrado && !isProvinciaField) {
                                const newValues = [...editedValues]
                                newValues[idx] = e.target.value
                                setEditedValues(newValues)
                              }
                            }}
                            disabled={isIdField || isAutoField || isCampoCerrado || isProvinciaField}
                            className={`${isAutoField ? 'auto-input' : ''} ${isCampoObligatorio ? 'input-obligatorio' : ''}`}
                            aria-required={isCampoObligatorio}
                          />
                        )}
                      </div>
                    )
                  })}
                </div>
                
                {/* Sección de Subida de Imagen */}
                <div className="image-upload-section">
                  <h3>📷 Foto del PDV</h3>
                  <div className="image-upload-container">
                    {imagePreview ? (
                      <div className="image-preview-wrapper">
                        <Image 
                          src={imagePreview} 
                          alt="Vista previa de la foto del PDV" 
                          className="image-preview"
                          width={400}
                          height={300}
                          unoptimized
                        />
                        <div className="image-actions">
                          {uploadedImageUrl && (
                            <a 
                              href={uploadedImageUrl} 
                              target="_blank" 
                              rel="noopener noreferrer"
                              className="btn-view-image"
                            >
                              🔗 Ver imagen completa
                            </a>
                          )}
                          <button 
                            type="button"
                            className="btn-remove-image"
                            onClick={handleClearImage}
                            disabled={uploadingImage}
                          >
                            ✕ Quitar imagen
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="image-upload-options">
                        {uploadingImage ? (
                          <div className="upload-loading">
                            <div className="spinner"></div>
                            <span>Subiendo imagen...</span>
                          </div>
                        ) : (
                          <>
                            {/* Opción 1: Galería */}
                            <label className="image-upload-option">
                        <input
                          type="file"
                          accept="image/*"
                          onChange={handleImageUpload}
                          disabled={uploadingImage}
                          className="image-input-hidden"
                        />
                              <div className="upload-option-content">
                                <span className="upload-option-icon">🖼️</span>
                                <span className="upload-option-text">Galería</span>
                              </div>
                            </label>
                            {/* Opción 2: Cámara (móvil) - Captura ubicación GPS - Solo visible en móvil */}
                            <label className="image-upload-option camera-option">
                              <input
                                type="file"
                                accept="image/*"
                                capture="environment"
                                onChange={(e) => handleImageUpload(e, true)}
                                disabled={uploadingImage}
                                className="image-input-hidden"
                              />
                              <div className="upload-option-content">
                                <span className="upload-option-icon">📷</span>
                                <span className="upload-option-text">Cámara</span>
                                <span className="upload-option-hint">📍 GPS</span>
                              </div>
                            </label>
                          </>
                        )}
                        <span className="upload-hint-bottom">JPG, PNG o GIF (máx. 32MB)</span>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="modal-footer">
                <button className="btn-secondary" onClick={handleCancelEdit} disabled={saving}>
                  Cancelar
                </button>
                <button className="btn-primary" onClick={() => handleSaveRow(false)} disabled={saving || uploadingImage}>
                  {saving ? 'Guardando...' : 'Guardar Cambios'}
                </button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Modal de Opciones de Ubicación */}
      {locationModalRow !== null && !showMapPicker && (
        <div className="modal-overlay" onClick={() => {
          setLocationModalRow(null)
          // Si vino de búsqueda mobile, volver a ella
          if (cameFromMobileSearch) {
            setShowMobileSearch(true)
            setCameFromMobileSearch(false)
          }
        }}>
          <div className="modal-content modal-small location-options" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>📍 Opciones de Ubicación</h2>
              <button className="modal-close" onClick={() => {
                setLocationModalRow(null)
                // Si vino de búsqueda mobile, volver a ella
                if (cameFromMobileSearch) {
                  setShowMobileSearch(true)
                  setCameFromMobileSearch(false)
                }
              }}>×</button>
            </div>
            <div className="modal-body">
              <p className="location-description">¿Cómo deseas agregar la ubicación?</p>
              
              {/* Mostrar coordenadas existentes si las hay */}
              {(() => {
                if (!sheetData?.data?.length || locationModalRow === null || locationModalRow < 0 || locationModalRow >= sheetData.data.length) return null
                const headers = sheetData.headers.map(h => h.toLowerCase().trim())
                const latIndex = headers.findIndex(h => h === 'latitud' || h === 'lat')
                const lngIndex = headers.findIndex(h => h === 'longitud' || h === 'lng' || h === 'long')
                const row = sheetData.data[locationModalRow]
                const existingLat = latIndex !== -1 && row ? String(row[latIndex] || '').trim() : ''
                const existingLng = lngIndex !== -1 && row ? String(row[lngIndex] || '').trim() : ''
                
                if (existingLat || existingLng) {
                  return (
                    <div className="existing-coords-info">
                      <strong>⚠️ Coordenadas actuales:</strong>
                      <p>Lat: {existingLat || '(vacío)'}</p>
                      <p>Lng: {existingLng || '(vacío)'}</p>
                      <small>Se sobrescribirán al guardar</small>
                    </div>
                  )
                }
                return null
              })()}
              
              <div className="location-buttons">
                <button 
                  className="btn-location-option btn-gps"
                  onClick={handleSaveCurrentLocation}
                  disabled={savingLocationRow !== null}
                >
                  {savingLocationRow !== null ? (
                    <>
                      <span className="btn-icon"><div className="saving-spinner-medium"></div></span>
                      <span className="btn-text">Obteniendo ubicación...</span>
                      <span className="btn-hint">Por favor espera</span>
                    </>
                  ) : (
                    <>
                      <span className="btn-icon">🛰️</span>
                      <span className="btn-text">Utilizar ubicación actual</span>
                      <span className="btn-hint">GPS automático</span>
                    </>
                  )}
                </button>
                
                <button 
                  className="btn-location-option btn-map"
                  onClick={handleOpenMapPicker}
                  disabled={savingLocationRow !== null}
                >
                  <span className="btn-icon">🗺️</span>
                  <span className="btn-text">Agregar ubicación manual</span>
                  <span className="btn-hint">Seleccionar en mapa</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Selector de Mapa */}
      {showMapPicker && manualCoords && (
        <div className="modal-overlay">
          <div className="modal-content modal-map" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>🗺️ Seleccionar Ubicación</h2>
              <button className="modal-close" onClick={() => {
                setShowMapPicker(false)
                setManualCoords(null)
                setShowLocationWarning(false)
              }}>×</button>
            </div>
            <div className="modal-body map-body">
              {/* Popup de advertencia para revisar ubicación - OBLIGATORIO */}
              {showLocationWarning && (
                <div className="location-warning-overlay">
                  <div className="location-warning-popup">
                    <div className="location-warning-content">
                      <span className="location-warning-icon">📍</span>
                      <div className="location-warning-text">
                        <strong>Por favor, revisar ubicación</strong>
                        <p>Verificá que el marcador esté en la posición correcta del puesto</p>
                      </div>
                    </div>
                    <button 
                      className="location-warning-close"
                      onClick={() => setShowLocationWarning(false)}
                    >
                      ✓ Entendido
                    </button>
                  </div>
                </div>
              )}
              
              <p className="map-instructions">📍 Busca una dirección o arrastra el marcador en el mapa</p>
              <p className="map-instructions-hint">(Busca una dirección aproximada para poder ubicarla con mayor precisión)</p>
              
              {/* Buscador de direcciones */}
              <div className="address-search">
                <input
                  type="text"
                  placeholder="Ej: Santa Fe 300, Palermo"
                  value={addressSearch}
                  onChange={(e) => setAddressSearch(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault()
                      handleSearchAddress()
                    }
                  }}
                  disabled={searchingAddress}
                />
                <button 
                  onClick={handleSearchAddress}
                  disabled={searchingAddress}
                  className="btn-search-address"
                >
                  {searchingAddress ? '🔄' : '🔍'} {searchingAddress ? 'Buscando...' : 'Buscar'}
                </button>
              </div>
              
              <div className="map-container">
                <MapPicker
                  initialLat={manualCoords.lat}
                  initialLng={manualCoords.lng}
                  onCoordsChange={(lat, lng) => setManualCoords({ lat, lng })}
                />
              </div>
              
              <p className="map-click-hint">👆 Haz clic en el mapa para establecer la ubicación exacta</p>
              
              <div className="coords-display">
                <span className="coord-label">📍 Coordenadas:</span>
                <span className="coord-value">{manualCoords.lat.toFixed(6)}, {manualCoords.lng.toFixed(6)}</span>
              </div>
              
              {/* Overlay de carga para búsqueda de dirección */}
              {searchingAddress && (
                <div className="searching-overlay">
                  <div className="searching-content">
                    <div className="searching-spinner"></div>
                    <p>Buscando dirección...</p>
                  </div>
                </div>
              )}
              
              {/* Overlay de carga para guardar */}
              {savingManualLocation && (
                <div className="saving-overlay">
                  <div className="saving-content">
                    <div className="saving-spinner-large"></div>
                    <p>Guardando ubicación...</p>
                    <span>Por favor espera</span>
                  </div>
                </div>
              )}
            </div>
            
            {/* Botones fuera del modal body - siempre visibles */}
            <div className="map-actions-fixed">
              <button 
                className="btn-cancel-red"
                onClick={() => {
                  const confirmed = window.confirm(
                    '¿Estás seguro de cancelar?\n\n' +
                    'Se perderán las coordenadas seleccionadas.'
                  )
                  if (confirmed) {
                    setShowMapPicker(false)
                    setManualCoords(null)
                    setShowLocationWarning(false)
                  }
                }}
                disabled={savingManualLocation}
              >
                ✕ Cancelar
              </button>
              <button 
                className="btn-accept-green"
                onClick={() => {
                  const confirmed = window.confirm(
                    `📍 ¿Guardar esta ubicación?\n\n` +
                    `Latitud: ${manualCoords?.lat.toFixed(6)}\n` +
                    `Longitud: ${manualCoords?.lng.toFixed(6)}`
                  )
                  if (confirmed) {
                    handleSaveManualLocation()
                  }
                }}
                disabled={savingManualLocation}
              >
                {savingManualLocation ? (
                  <>
                    <span className="saving-spinner"></span>
                    Guardando...
                  </>
                ) : (
                  '✓ Aceptar'
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Opciones Ayuda / Soporte */}
      {showNuevoPdvModal && (
        <div className="modal-overlay" onClick={() => setShowNuevoPdvModal(false)}>
          <div className="modal-content modal-small nuevo-pdv-options" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>⚠️ Ayuda / Soporte</h2>
              <button className="modal-close" onClick={() => setShowNuevoPdvModal(false)}>×</button>
            </div>
            <div className="modal-body">
              <p className="nuevo-pdv-description">Selecciona una opción:</p>
              <div className="nuevo-pdv-buttons">
                {/* Opción Agregar nuevo PDV - Solo visible para admins */}
                {sheetData?.permissions?.isAdmin && (
                  <button 
                    className="nuevo-pdv-option-btn option-agregar"
                    onClick={() => {
                      setShowNuevoPdvModal(false)
                      setShowNuevoPdvForm(true)
                      resetNuevoPdvForm()
                    }}
                  >
                    <span className="option-icon">🏪</span>
                    <span className="option-title">AGREGAR NUEVO PDV</span>
                    <span className="option-desc">Registrar un nuevo punto de venta en el sistema</span>
                  </button>
                )}
                {/* Opción Cuestionario PDF - Visible para todos */}
                <button 
                  className="nuevo-pdv-option-btn option-cuestionario"
                  onClick={() => {
                    setShowNuevoPdvModal(false)
                    downloadCuestionario()
                  }}
                >
                  <span className="option-icon">📋</span>
                  <span className="option-title">CUESTIONARIO PDF</span>
                  <span className="option-desc">Descargar formulario para llenar a mano</span>
                </button>
                {/* Opción Soporte WhatsApp - Visible para todos */}
                <a 
                  href="https://api.whatsapp.com/send/?phone=541165903360&text&type=phone_number&app_absent=0"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="nuevo-pdv-option-btn option-whatsapp"
                  onClick={() => setShowNuevoPdvModal(false)}
                >
                  <span className="option-icon">💬</span>
                  <span className="option-title">SOPORTE WHATSAPP</span>
                  <span className="option-desc">Contactar soporte técnico por WhatsApp</span>
                </a>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal Formulario Nuevo PDV */}
      {showNuevoPdvForm && (
        <div className="modal-overlay" onClick={() => !savingNuevoPdv && setShowNuevoPdvForm(false)}>
          <div className="modal-content modal-large" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>🏪 Agregar Nuevo PDV</h2>
              <button className="modal-close" onClick={() => !savingNuevoPdv && setShowNuevoPdvForm(false)} disabled={savingNuevoPdv}>×</button>
            </div>
            <div className="modal-body">
              <div className="nuevo-pdv-notice">
                <span className="notice-icon">ℹ️</span>
                <span>Este PDV se guardará en la hoja "ALTA PDV" con un nuevo ID asignado automáticamente.</span>
              </div>
              
              <div className="edit-form nuevo-pdv-form">
                {/* 1. Estado Kiosco */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-estadoKiosco">Estado del Kiosco</label>
                  <select
                    id="nuevo-pdv-estadoKiosco"
                    value={nuevoPdvData.estadoKiosco}
                    onChange={(e) => {
                      const nuevoEstado = e.target.value
                      setNuevoPdvData({
                        ...nuevoPdvData,
                        estadoKiosco: nuevoEstado,
                        reparto: nuevoEstado === 'Cerrado pero hace reparto' ? 'Si' : nuevoPdvData.reparto
                      })
                    }}
                    disabled={savingNuevoPdv}
                  >
                    {nuevoPdvOptions.estadoKiosco.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 2. Paquete */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-paquete">Paquete <span className="required">*</span></label>
                  <input
                    id="nuevo-pdv-paquete"
                    type="text"
                    value={nuevoPdvData.paquete}
                    onChange={(e) => {
                      setNuevoPdvData({...nuevoPdvData, paquete: e.target.value})
                      if (nuevoPdvFieldErrors.paquete) setNuevoPdvFieldErrors(prev => ({ ...prev, paquete: '' }))
                    }}
                    placeholder="Nombre del paquete"
                    disabled={savingNuevoPdv}
                    aria-required
                    aria-invalid={!!nuevoPdvFieldErrors.paquete}
                  />
                  {nuevoPdvFieldErrors.paquete && <span className="field-error" role="alert">{nuevoPdvFieldErrors.paquete}</span>}
                </div>

                {/* 3. Domicilio */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-domicilio">Domicilio completo <span className="required">*</span></label>
                  <input
                    id="nuevo-pdv-domicilio"
                    type="text"
                    value={nuevoPdvData.domicilio}
                    onChange={(e) => {
                      setNuevoPdvData({...nuevoPdvData, domicilio: e.target.value})
                      if (nuevoPdvFieldErrors.domicilio) setNuevoPdvFieldErrors(prev => ({ ...prev, domicilio: '' }))
                    }}
                    placeholder="Calle, número, esquina, etc."
                    disabled={savingNuevoPdv}
                    aria-required
                    aria-invalid={!!nuevoPdvFieldErrors.domicilio}
                  />
                  {nuevoPdvFieldErrors.domicilio && <span className="field-error" role="alert">{nuevoPdvFieldErrors.domicilio}</span>}
                </div>

                {/* 4. Provincia */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-provincia">Provincia</label>
                  <input
                    id="nuevo-pdv-provincia"
                    type="text"
                    value={nuevoPdvData.provincia}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, provincia: e.target.value})}
                    disabled={savingNuevoPdv}
                  />
                </div>

                {/* 5. Partido */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-partido">Partido</label>
                  <select
                    id="nuevo-pdv-partido"
                    value={nuevoPdvData.partido}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, partido: e.target.value, localidad: ''})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar partido...</option>
                    {nuevoPdvOptions.partido.map(p => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                </div>

                {/* 6. Localidad / Barrio - Campo predictivo */}
                <div className="edit-field autocomplete-field">
                  <label htmlFor="nuevo-pdv-localidad">Localidad / Barrio</label>
                  <div className="autocomplete-container">
                    <input
                      id="nuevo-pdv-localidad"
                      type="text"
                      value={nuevoPdvData.localidad}
                      onChange={(e) => handleLocalidadChange(e.target.value)}
                      onFocus={() => {
                        if (nuevoPdvData.partido && nuevoPdvData.localidad.length > 0) {
                          const localidades = (nuevoPdvOptions.localidadesPorPartido as Record<string, string[]>)[nuevoPdvData.partido] || []
                          const filtered = localidades.filter(loc => 
                            loc.toLowerCase().includes(nuevoPdvData.localidad.toLowerCase())
                          )
                          setLocalidadSugerencias(filtered)
                          setShowLocalidadSugerencias(filtered.length > 0)
                        } else if (nuevoPdvData.partido) {
                          setLocalidadSugerencias((nuevoPdvOptions.localidadesPorPartido as Record<string, string[]>)[nuevoPdvData.partido] || [])
                          setShowLocalidadSugerencias(true)
                        }
                      }}
                      onBlur={() => setTimeout(() => setShowLocalidadSugerencias(false), 200)}
                      placeholder={nuevoPdvData.partido ? 'Escriba para buscar localidad...' : 'Primero seleccione un partido'}
                      disabled={savingNuevoPdv || !nuevoPdvData.partido}
                      autoComplete="off"
                    />
                    {showLocalidadSugerencias && localidadSugerencias.length > 0 && (
                      <ul className="autocomplete-suggestions">
                        {localidadSugerencias.map((loc, idx) => (
                          <li 
                            key={idx} 
                            onMouseDown={() => selectLocalidad(loc)}
                            className="autocomplete-item"
                          >
                            {loc}
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>

                {/* 7. N° Vendedor */}
                <div className="edit-field">
                  <label>N° Vendedor</label>
                  <input
                    type="text"
                    value={nuevoPdvData.nVendedor}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, nVendedor: e.target.value})}
                    disabled={savingNuevoPdv}
                  />
                </div>

                {/* 8. Distribuidora */}
                <div className="edit-field">
                  <label>Distribuidora</label>
                  <select
                    value={nuevoPdvData.distribuidora}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, distribuidora: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.distribuidora.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 9. Días de atención */}
                <div className="edit-field">
                  <label>Días de atención</label>
                  <select
                    value={nuevoPdvData.diasAtencion}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, diasAtencion: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.diasAtencion.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 10. Horario */}
                <div className="edit-field">
                  <label>Horario</label>
                  <select
                    value={nuevoPdvData.horario}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, horario: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.horario.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 11. Escaparate */}
                <div className="edit-field">
                  <label>Escaparate</label>
                  <select
                    value={nuevoPdvData.escaparate}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, escaparate: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.escaparate.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 12. Ubicación */}
                <div className="edit-field">
                  <label>Ubicación</label>
                  <select
                    value={nuevoPdvData.ubicacion}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, ubicacion: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.ubicacion.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 13. Fachada del puesto */}
                <div className="edit-field">
                  <label>Fachada del puesto</label>
                  <select
                    value={nuevoPdvData.fachada}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, fachada: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.fachada.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 14. Venta productos no editoriales */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-ventaNoEditorial">Venta prod. no editoriales <span className="required">*</span></label>
                  <select
                    id="nuevo-pdv-ventaNoEditorial"
                    value={nuevoPdvData.ventaNoEditorial}
                    onChange={(e) => {
                      setNuevoPdvData({...nuevoPdvData, ventaNoEditorial: e.target.value})
                      if (nuevoPdvFieldErrors.ventaNoEditorial) setNuevoPdvFieldErrors(prev => ({ ...prev, ventaNoEditorial: '' }))
                    }}
                    disabled={savingNuevoPdv}
                    aria-required
                    aria-invalid={!!nuevoPdvFieldErrors.ventaNoEditorial}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.ventaNoEditorial.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                  {nuevoPdvFieldErrors.ventaNoEditorial && <span className="field-error" role="alert">{nuevoPdvFieldErrors.ventaNoEditorial}</span>}
                </div>

                {/* 15. Reparto */}
                <div className="edit-field">
                  <label>Reparto</label>
                  <select
                    value={nuevoPdvData.reparto}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, reparto: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.reparto.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 16. Suscripciones */}
                <div className="edit-field">
                  <label>Suscripciones</label>
                  <select
                    value={nuevoPdvData.suscripciones}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, suscripciones: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.suscripciones.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 17. Nombre y Apellido */}
                <div className="edit-field">
                  <label>Nombre y Apellido</label>
                  <input
                    type="text"
                    value={nuevoPdvData.nombreApellido}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, nombreApellido: e.target.value})}
                    placeholder="Nombre del vendedor/encargado"
                    disabled={savingNuevoPdv}
                  />
                </div>

                {/* 18. Mayor venta */}
                <div className="edit-field">
                  <label>Mayor venta</label>
                  <select
                    value={nuevoPdvData.mayorVenta}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, mayorVenta: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.mayorVenta.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 19. Utiliza Parada Online */}
                <div className="edit-field">
                  <label>¿Utiliza Parada Online?</label>
                  <select
                    value={nuevoPdvData.paradaOnline}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, paradaOnline: e.target.value})}
                    disabled={savingNuevoPdv}
                  >
                    <option value="">Seleccionar...</option>
                    {nuevoPdvOptions.paradaOnline.map(opt => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>

                {/* 20. Teléfono */}
                <div className="edit-field">
                  <label htmlFor="nuevo-pdv-telefono">Teléfono <span className="required">*</span></label>
                  <input
                    id="nuevo-pdv-telefono"
                    type="text"
                    value={nuevoPdvData.telefono}
                    onChange={(e) => {
                      setNuevoPdvData({...nuevoPdvData, telefono: e.target.value})
                      if (nuevoPdvFieldErrors.telefono) setNuevoPdvFieldErrors(prev => ({ ...prev, telefono: '' }))
                    }}
                    placeholder="Poner 0 si no se obtiene"
                    disabled={savingNuevoPdv}
                    aria-required
                    aria-invalid={!!nuevoPdvFieldErrors.telefono}
                  />
                  {nuevoPdvFieldErrors.telefono && <span className="field-error" role="alert">{nuevoPdvFieldErrors.telefono}</span>}
                </div>

                {/* 21. Correo electrónico */}
                <div className="edit-field">
                  <label>Correo electrónico</label>
                  <input
                    type="email"
                    value={nuevoPdvData.correoElectronico}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, correoElectronico: e.target.value})}
                    placeholder="email@ejemplo.com"
                    disabled={savingNuevoPdv}
                  />
                </div>

                {/* 22. Observaciones */}
                <div className="edit-field">
                  <label>Observaciones</label>
                  <textarea
                    value={nuevoPdvData.observaciones}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, observaciones: e.target.value})}
                    placeholder="Observaciones del PDV..."
                    rows={2}
                    disabled={savingNuevoPdv}
                  />
                </div>

                {/* 23. Comentarios */}
                <div className="edit-field">
                  <label>Comentarios</label>
                  <textarea
                    value={nuevoPdvData.comentarios}
                    onChange={(e) => setNuevoPdvData({...nuevoPdvData, comentarios: e.target.value})}
                    placeholder="Comentarios adicionales..."
                    rows={2}
                    disabled={savingNuevoPdv}
                  />
                </div>

                {/* Imagen */}
                <div className="image-upload-section">
                  <h3>📷 Foto del PDV</h3>
                  <div className="image-upload-container">
                    {nuevoPdvImagePreview ? (
                      <div className="image-preview-wrapper">
                        <Image src={nuevoPdvImagePreview} alt="Vista previa de la foto del PDV" className="image-preview" width={400} height={300} unoptimized />
                        <div className="image-actions">
                          {nuevoPdvImageUrl && (
                            <a href={nuevoPdvImageUrl} target="_blank" rel="noopener noreferrer" className="btn-view-image">
                              🔗 Ver imagen
                            </a>
                          )}
                          <button 
                            type="button"
                            className="btn-remove-image"
                            onClick={() => {
                              setNuevoPdvImagePreview(null)
                              setNuevoPdvImageUrl(null)
                            }}
                            disabled={uploadingNuevoPdvImage || savingNuevoPdv}
                          >
                            ✕ Quitar
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="image-upload-options">
                        {uploadingNuevoPdvImage ? (
                          <div className="upload-loading">
                            <div className="spinner"></div>
                            <span>Subiendo imagen...</span>
                          </div>
                        ) : (
                          <>
                            {/* Opción 1: Galería */}
                            <label className="image-upload-option">
                              <input
                                type="file"
                                accept="image/*"
                                onChange={handleNuevoPdvImageUpload}
                                disabled={uploadingNuevoPdvImage || savingNuevoPdv}
                                className="image-input-hidden"
                              />
                              <div className="upload-option-content">
                                <span className="upload-option-icon">🖼️</span>
                                <span className="upload-option-text">Galería</span>
                              </div>
                            </label>
                            {/* Opción 2: Cámara (móvil) - Captura ubicación GPS - Solo visible en móvil */}
                            <label className="image-upload-option camera-option">
                              <input
                                type="file"
                                accept="image/*"
                                capture="environment"
                                onChange={(e) => handleNuevoPdvImageUpload(e, true)}
                                disabled={uploadingNuevoPdvImage || savingNuevoPdv}
                                className="image-input-hidden"
                              />
                              <div className="upload-option-content">
                                <span className="upload-option-icon">📷</span>
                                <span className="upload-option-text">Cámara</span>
                                <span className="upload-option-hint">📍 GPS</span>
                              </div>
                            </label>
                          </>
                        )}
                        <span className="upload-hint-bottom">JPG, PNG o GIF (máx. 32MB)</span>
                      </div>
                    )}
                  </div>
                  </div>
                </div>
              </div>
              <div className="modal-footer">
              <button 
                className="btn-secondary" 
                onClick={() => setShowNuevoPdvForm(false)} 
                disabled={savingNuevoPdv}
              >
                  Cancelar
                </button>
              <button 
                className="btn-primary" 
                onClick={handleSaveNuevoPdv} 
                disabled={savingNuevoPdv || uploadingNuevoPdvImage}
              >
                {savingNuevoPdv ? 'Guardando...' : '✓ Dar de Alta PDV'}
                </button>
              </div>
            </div>
          </div>
      )}

      {/* Admin Sidebar - visible para admin y supervisor */}
      {showAdminSidebar && (sheetData?.permissions?.role === 'admin' || sheetData?.permissions?.role === 'supervisor') && (
        <>
          <div className="sidebar-overlay" onClick={() => setShowAdminSidebar(false)}></div>
          <aside className="admin-sidebar">
            <div className="sidebar-header">
              <h2>⚙️ Panel Admin</h2>
              <button className="sidebar-close" onClick={() => setShowAdminSidebar(false)}>×</button>
            </div>
            
            <nav className="sidebar-nav">
              <Link 
                href="/mapa" 
                className="sidebar-nav-btn sidebar-nav-link"
                onClick={() => setShowAdminSidebar(false)}
              >
                🗺️ Ver mapa
              </Link>
              <button 
                className="sidebar-nav-btn sidebar-nav-action"
                onClick={() => {
                  setShowAdminSidebar(false)
                  setShowNuevoPdvModal(true)
                }}
              >
                ➕ Nuevo PDV
              </button>
              <button 
                className={`sidebar-nav-btn ${adminSidebarTab === 'hojas' ? 'active' : ''}`}
                onClick={() => setAdminSidebarTab('hojas')}
              >
                📋 Hojas
              </button>
              <button 
                className={`sidebar-nav-btn ${adminSidebarTab === 'usuarios' ? 'active' : ''}`}
                onClick={() => setAdminSidebarTab('usuarios')}
              >
                👥 Usuarios
              </button>
              <button 
                className={`sidebar-nav-btn ${adminSidebarTab === 'stats' ? 'active' : ''}`}
                onClick={() => setAdminSidebarTab('stats')}
              >
                📊 Estadísticas
              </button>
              <button 
                className={`sidebar-nav-btn ${adminSidebarTab === 'reportes' ? 'active' : ''}`}
                onClick={() => setAdminSidebarTab('reportes')}
              >
                📥 Reportes
              </button>
              {/* Seguimiento GPS - Solo visible para admins (no supervisores) */}
              {sheetData?.permissions?.role === 'admin' && (
                <button 
                  className={`sidebar-nav-btn ${adminSidebarTab === 'gps' ? 'active' : ''}`}
                  onClick={() => {
                    setAdminSidebarTab('gps')
                    loadGpsLogs()
                  }}
                >
                  📍 Seguimiento GPS
                </button>
              )}
            </nav>
            
            <div className="sidebar-content">
              {/* Pestaña Hojas */}
              {adminSidebarTab === 'hojas' && (
                <div className="sidebar-section">
                  <h3>Seleccionar Hoja</h3>
                  <div className="sidebar-sheet-list">
                    <button
                      className={`sidebar-sheet-btn ${adminSelectedSheet === 'Todos' || adminSelectedSheet === '' ? 'active' : ''}`}
                      onClick={() => {
                        setAdminSelectedSheet('Todos')
                        setCurrentPage(1)
                        if (accessToken) loadSheetData(accessToken, 'Todos')
                      }}
                    >
                      📊 Todos
                    </button>
                    {availableSheets
                      // Ocultar hoja "test" para supervisores (solo admins pueden verla)
                      .filter(sheet => sheetData?.permissions?.role === 'admin' || sheet.toLowerCase() !== 'test')
                      .map((sheet, idx) => (
                        <button
                          key={idx}
                          className={`sidebar-sheet-btn ${adminSelectedSheet === sheet ? 'active' : ''}`}
                          onClick={() => {
                            setAdminSelectedSheet(sheet)
                            setCurrentPage(1)
                            if (accessToken) loadSheetData(accessToken, sheet)
                          }}
                        >
                          📋 {sheet}
                        </button>
                      ))}
                  </div>
                  
                  {/* Hojas asignadas a usuarios - útil para supervisores */}
                  {allPermissions
                    .filter(p => p.assignedSheet)
                    // Ocultar usuarios con hoja "test" para supervisores
                    .filter(p => sheetData?.permissions?.role === 'admin' || p.assignedSheet?.toLowerCase() !== 'test')
                    .length > 0 && (
                    <div className="sidebar-assigned-sheets">
                      <h4>👥 Hojas por Usuario</h4>
                      <div className="assigned-sheets-list">
                        {allPermissions
                          .filter(p => p.assignedSheet)
                          // Ocultar usuarios con hoja "test" para supervisores
                          .filter(p => sheetData?.permissions?.role === 'admin' || p.assignedSheet?.toLowerCase() !== 'test')
                          .map((perm, idx) => {
                            const userStats = getUserStats(perm.email, perm.allowedIds)
                            const percent = userStats.total > 0 ? Math.round((userStats.relevados / userStats.total) * 100) : 0
                            return (
                              <button
                                key={idx}
                                className={`assigned-sheet-btn ${adminSelectedSheet === perm.assignedSheet ? 'active' : ''}`}
                                onClick={() => {
                                  const sheet = perm.assignedSheet ?? ''
                                  if (!sheet) return
                                  setAdminSelectedSheet(sheet)
                                  setCurrentPage(1)
                                  if (accessToken) loadSheetData(accessToken, sheet)
                                }}
                              >
                                <span className="assigned-user-name">{perm.email.split('@')[0]}</span>
                                <span className="assigned-sheet-name">📋 {perm.assignedSheet}</span>
                                <span className="assigned-progress">{percent}%</span>
                              </button>
                            )
                          })}
                      </div>
                    </div>
                  )}
                  
                  <div className="sidebar-sheet-info">
                    <span className="info-label">Hoja actual:</span>
                    <span className="info-value">{adminSelectedSheet || 'Hoja 1'}</span>
                  </div>
                  <div className="sidebar-sheet-info">
                    <span className="info-label">Total registros:</span>
                    <span className="info-value">{sheetData?.data?.length || 0}</span>
                  </div>
                </div>
              )}
              
              {/* Pestaña Usuarios */}
              {adminSidebarTab === 'usuarios' && (
                <div className="sidebar-section">
                  <h3>Gestión de Usuarios</h3>
                  <p className="sidebar-description">Administra los permisos de acceso de los usuarios.</p>
                  <button 
                    className="sidebar-action-btn"
                    onClick={() => {
                      setShowAdminSidebar(false)
                      setShowPermissionsPanel(true)
                    }}
                  >
                    👥 Abrir Panel de Usuarios
                  </button>
                  <div className="sidebar-user-summary">
                    <span className="summary-item">
                      <strong>{allPermissions.length}</strong> usuarios configurados
                    </span>
                  </div>
                </div>
              )}
              
              {/* Pestaña Estadísticas */}
              {adminSidebarTab === 'stats' && (
                <div className="sidebar-section">
                  <h3>Estadísticas</h3>
                  <p className="sidebar-description">Visualiza el progreso del relevamiento por hoja.</p>
                  <button 
                    className="sidebar-action-btn"
                    onClick={() => {
                      setShowAdminSidebar(false)
                      setShowStats(true)
                      setStatsSelectedSheet('Total')
                      loadAllSheetsStats()
                    }}
                  >
                    📊 Ver Estadísticas Completas
                  </button>
                  
                  {/* Mini resumen de stats */}
                  {Object.keys(allSheetsStats).length > 0 && (
                    <div className="sidebar-stats-mini">
                      {Object.keys(allSheetsStats)
                        .filter(name => name !== 'Hoja 1')
                        .map(sheetName => {
                          const data = allSheetsStats[sheetName]
                          if (!data || data.length <= 1) return null
                          const headers = data[0] as string[]
                          const rows = data.slice(1)
                          const relevadorIdx = headers.findIndex(h => 
                            String(h).toLowerCase().includes('relevado por')
                          )
                          const relevados = relevadorIdx !== -1 
                            ? rows.filter(r => String(r[relevadorIdx] || '').trim() !== '').length 
                            : 0
                          const total = rows.length
                          const percent = total > 0 ? Math.round((relevados / total) * 100) : 0
                          
                          return (
                            <div key={sheetName} className="mini-stat-item">
                              <span className="mini-stat-name">{sheetName}</span>
                              <div className="mini-stat-bar">
                                <div className="mini-stat-fill" style={{ width: `${percent}%` }}></div>
                              </div>
                              <span className="mini-stat-percent">{percent}%</span>
                            </div>
                          )
                        })}
                    </div>
                  )}
                </div>
              )}
              
              {/* Pestaña Reportes */}
              {adminSidebarTab === 'reportes' && (
                <div className="sidebar-section">
                  <h3>Descargar Reportes</h3>
                  <p className="sidebar-description">Exporta los datos en formato Excel/CSV.</p>
                  
                  <div className="sidebar-report-options">
                    <button 
                      className="sidebar-report-btn"
                      onClick={() => downloadSheetReport()}
                      disabled={!sheetData}
                    >
                      📥 Descargar Hoja Actual
                      <span className="report-hint">{adminSelectedSheet || 'Hoja 1'}</span>
                    </button>
                    
                    <button 
                      className="sidebar-report-btn"
                      onClick={() => {
                        setShowAdminSidebar(false)
                        setShowStats(true)
                        setStatsSelectedSheet('Total')
                        loadAllSheetsStats()
                      }}
                    >
                      📊 Descargar Estadísticas
                      <span className="report-hint">Desde el panel de stats</span>
                    </button>
                  </div>
                </div>
              )}
              
              {/* Pestaña Seguimiento GPS - Solo visible para admins (no supervisores) */}
              {adminSidebarTab === 'gps' && sheetData?.permissions?.role === 'admin' && (
                <div className="sidebar-section">
                  <h3>📍 Seguimiento GPS</h3>
                  <p className="sidebar-description">Ver ubicaciones de usuarios móviles.</p>
                  
                  {loadingGpsLogs ? (
                    <div className="sidebar-loading">Cargando logs GPS...</div>
                  ) : (
                    <>
                      <div className="gps-user-select">
                        <label>Seleccionar usuario:</label>
                        <select 
                          value={selectedGpsUser}
                          onChange={(e) => {
                            setSelectedGpsUser(e.target.value)
                            loadGpsLogs(e.target.value || undefined)
                          }}
                        >
                          <option value="">-- Todos los usuarios --</option>
                          {gpsUsers.map((user, idx) => (
                            <option key={idx} value={user}>{user}</option>
                          ))}
                        </select>
                      </div>
                      
                      <div className="gps-stats">
                        <div className="gps-stat">
                          <span className="stat-number">{gpsLogs.length}</span>
                          <span className="stat-label">Registros</span>
                        </div>
                        <div className="gps-stat">
                          <span className="stat-number">{gpsUsers.length}</span>
                          <span className="stat-label">Usuarios</span>
                        </div>
                      </div>
                      
                      <button 
                        className="btn-view-map"
                        onClick={() => setShowGpsModal(true)}
                        disabled={gpsLogs.length === 0}
                      >
                        🗺️ Ver en Mapa
                      </button>
                      
                      {gpsLogs.length > 0 && (
                        <div className="gps-recent-list">
                          <h4>Últimos registros:</h4>
                          {gpsLogs.slice(0, 5).map((log, idx) => (
                            <div key={idx} className="gps-log-item">
                              <div className="log-email">{log.email}</div>
                              <div className="log-details">
                                <span>{log.fecha} {log.hora}</span>
                                <span className="log-event">{log.evento}</span>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </>
                  )}
                </div>
              )}
            </div>
          </aside>
        </>
      )}

      {/* Modal de Mapa GPS (Solo Admin - supervisores NO tienen acceso) */}
      {showGpsModal && sheetData?.permissions?.role === 'admin' && (
        <div className="modal-overlay" onClick={() => setShowGpsModal(false)}>
          <div className="modal-content modal-gps-map" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>📍 Seguimiento GPS {selectedGpsUser && `- ${selectedGpsUser}`}</h2>
              <button className="modal-close" onClick={() => setShowGpsModal(false)}>×</button>
            </div>
            <div className="modal-body gps-map-body">
              <div className="gps-map-legend">
                <div className="legend-item">
                  <span className="legend-dot green"></span>
                  <span>Primer punto (más antiguo)</span>
                </div>
                <div className="legend-item">
                  <span className="legend-dot blue"></span>
                  <span>Puntos intermedios</span>
                </div>
                <div className="legend-item">
                  <span className="legend-dot red"></span>
                  <span>Último punto (más reciente)</span>
                </div>
              </div>
              
              <div className="gps-map-container">
                <GpsTrackingMap 
                  logs={gpsLogs}
                  selectedUser={selectedGpsUser}
                />
              </div>
              
              <div className="gps-map-info">
                <span>Total de puntos: <strong>{gpsLogs.length}</strong></span>
                {selectedGpsUser && <span>Usuario: <strong>{selectedGpsUser}</strong></span>}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Admin Panel Modal */}
      {showPermissionsPanel && sheetData?.permissions?.isAdmin && (
        <div className="modal-overlay" onClick={() => setShowPermissionsPanel(false)}>
          <div className="admin-panel" onClick={e => e.stopPropagation()}>
            <div className="admin-header">
              <h2>Panel de Administración</h2>
              <div className="admin-header-actions">
                <button 
                  className="btn-primary"
                  onClick={async () => {
                    if (accessToken) {
                      await loadSheetData(accessToken, adminSelectedSheet || 'Todos')
                      await loadPermissions(accessToken)
                    }
                  }}
                  disabled={loadingData}
                >
                  {loadingData ? 'Actualizando...' : 'Actualizar'}
                </button>
                <button className="btn-close" onClick={() => setShowPermissionsPanel(false)}>×</button>
              </div>
            </div>
            
            <div className="admin-body">
              <div className="admin-stats">
                <span>Total de registros: <strong>{sheetData?.data?.length ?? 0}</strong></span>
              </div>

              {/* Add new user form - Solo visible para admins (no supervisores) */}
              {sheetData?.permissions?.role === 'admin' && (
                <div className="add-user-form">
                  <input
                    type="email"
                    placeholder="Email del nuevo usuario"
                    value={newUserEmail}
                    onChange={(e) => setNewUserEmail(e.target.value)}
                    className="add-user-input"
                  />
                  <button 
                    className="btn-primary"
                    onClick={() => {
                      if (newUserEmail.trim()) {
                        setEditingPermission({ email: newUserEmail.trim(), allowedIds: [] })
                        setNewPermIds('')
                      }
                    }}
                    disabled={!newUserEmail.trim()}
                  >
                    Agregar Usuario
                  </button>
                </div>
              )}
              
              {/* Mensaje para supervisores */}
              {sheetData?.permissions?.role === 'supervisor' && (
                <div className="supervisor-notice">
                  <span className="notice-icon">👁️</span>
                  <span>Vista de solo lectura - Solo los administradores pueden modificar permisos</span>
                </div>
              )}

              {/* Filtro por rol */}
              <div className="users-role-filter">
                <label className="users-filter-label">Filtrar por rol:</label>
                <select
                  value={userRoleFilter}
                  onChange={(e) => setUserRoleFilter((e.target.value || 'all') as 'all' | '1' | '2' | '3')}
                  className="users-filter-select"
                >
                  <option value="all">Todos</option>
                  <option value="1">👤 Usuario</option>
                  <option value="2">👁️ Supervisor</option>
                  <option value="3">👑 Admin</option>
                </select>
              </div>

              <div className="users-grid">
                {allPermissions
                  .filter(perm => userRoleFilter === 'all' || (perm.level ?? 1) === Number(userRoleFilter))
                  .map((perm) => {
                    const originalIdx = allPermissions.indexOf(perm)
                    return (
                      <UserCard
                        key={originalIdx}
                        perm={perm}
                        originalIdx={originalIdx}
                        isAdmin={sheetData?.permissions?.role === 'admin'}
                        expandedIdsIndex={expandedIdsIndex}
                        onAssignIds={(p) => {
                          setEditingPermission(p)
                          setNewPermIds(p.allowedIds.join(', '))
                        }}
                        onToggleExpand={(idx) => setExpandedIdsIndex(expandedIdsIndex === idx ? null : idx)}
                      />
                    )
                  })}

                {allPermissions.length === 0 && (
                  <div className="no-users">
                    <p>No hay usuarios con permisos configurados</p>
                    <p className="hint">Agrega un usuario usando el formulario de arriba</p>
                  </div>
                )}
                {allPermissions.length > 0 && allPermissions.filter(perm => userRoleFilter === 'all' || (perm.level ?? 1) === Number(userRoleFilter)).length === 0 && (
                  <div className="no-users">
                    <p>No hay usuarios con el rol seleccionado</p>
                    <p className="hint">Prueba con otro filtro o &quot;Todos&quot;</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Statistics Modal */}
      {showStats && sheetData && (
        <div className="modal-overlay" onClick={() => setShowStats(false)}>
          <div className="stats-panel stats-panel-large" onClick={e => e.stopPropagation()}>
            <div className="stats-header">
              <h2>📊 Estadísticas de PDV Relevados</h2>
              <div className="stats-header-actions">
                <button 
                  className="btn-download-stats"
                  onClick={downloadStatsAsExcel}
                  title="Descargar estadísticas en Excel/CSV"
                >
                  📥 Descargar Excel
                </button>
                <button className="btn-close" onClick={() => setShowStats(false)}>×</button>
              </div>
            </div>
            
            {/* Tabs de hojas */}
            {sheetData.permissions?.isAdmin && Object.keys(allSheetsStats).length > 0 && (
              <div className="stats-tabs">
                <button
                  className={`stats-tab ${statsSelectedSheet === 'Total' ? 'active' : ''}`}
                  onClick={() => setStatsSelectedSheet('Total')}
                >
                  📊 Total
                </button>
                {Object.keys(allSheetsStats)
                  .filter(name => name !== 'Hoja 1') // Ocultar Hoja 1
                  .sort((a, b) => a.localeCompare(b))
                  .map(sheetName => (
                    <button
                      key={sheetName}
                      className={`stats-tab ${statsSelectedSheet === sheetName ? 'active' : ''}`}
                      onClick={() => setStatsSelectedSheet(sheetName)}
                    >
                      📋 {sheetName}
                    </button>
                  ))
                }
              </div>
            )}
            
            <div className="stats-body">
              {loadingStats ? (
                <div className="stats-loading">
                  <div className="loading-spinner"></div>
                  <p>Cargando estadísticas de todas las hojas...</p>
                </div>
              ) : sheetData.permissions?.isAdmin && Object.keys(allSheetsStats).length > 0 ? (
                // Vista de admin con estadísticas por hoja
                (() => {
                  const { stats, total, relevados, pendientes, relevadosConCoordenadas } = getFieldStatsForSheet(statsSelectedSheet)
                  
                  return (
                    <>
                      <div className="stats-sheet-title">
                        {statsSelectedSheet === 'Total' 
                          ? '📊 Estadísticas combinadas de todas las hojas'
                          : `📋 Estadísticas de: ${statsSelectedSheet}`
                        }
                      </div>
                      
                      <div className="stats-summary">
                        <div className="summary-card">
                          <span className="summary-number">{total}</span>
                          <span className="summary-label">Total PDV</span>
                        </div>
                        <div className="summary-card summary-green">
                          <span className="summary-number">{relevados}</span>
                          <span className="summary-label">Relevados</span>
                        </div>
                        <div className="summary-card summary-purple">
                          <span className="summary-number">{relevadosConCoordenadas ?? 0}</span>
                          <span className="summary-label">Relevados con Coordenadas</span>
                        </div>
                        <div className="summary-card summary-orange">
                          <span className="summary-number">{pendientes}</span>
                          <span className="summary-label">Pendientes</span>
                        </div>
                        <div className="summary-card summary-blue">
                          <span className="summary-number">
                            {total > 0 ? Math.round((relevados / total) * 100) : 0}%
                          </span>
                          <span className="summary-label">Progreso</span>
                        </div>
                      </div>
                      
                      <div className="charts-grid">
                        {stats.map((fieldStat, idx) => {
                          const fieldTotal = fieldStat.data.reduce((sum, d) => sum + d.count, 0)
                          let cumulativePercent = 0
                          
                          const gradientStops = fieldStat.data.map(d => {
                            const percent = (d.count / fieldTotal) * 100
                            const start = cumulativePercent
                            cumulativePercent += percent
                            return `${d.color} ${start}% ${cumulativePercent}%`
                          }).join(', ')
                          
                          return (
                            <div key={idx} className="chart-card">
                              <h3 className="chart-title">{fieldStat.fieldName}</h3>
                              <div className="chart-content">
                                <div 
                                  className="pie-chart"
                                  style={{ 
                                    background: `conic-gradient(${gradientStops})`
                                  }}
                                >
                                  <div className="pie-center">
                                    <span className="pie-total">{fieldTotal}</span>
                                    <span className="pie-label">Total</span>
                                  </div>
                                </div>
                                <div className="chart-legend">
                                  {fieldStat.data.slice(0, 8).map((d, i) => (
                                    <div key={i} className="legend-item">
                                      <span 
                                        className="legend-color" 
                                        style={{ background: d.color }}
                                      ></span>
                                      <span className="legend-label">{d.label}</span>
                                      <span className="legend-count">{d.count}</span>
                                      <span className="legend-percent">
                                        ({Math.round((d.count / fieldTotal) * 100)}%)
                                      </span>
                                    </div>
                                  ))}
                                  {fieldStat.data.length > 8 && (
                                    <div className="legend-more">
                                      +{fieldStat.data.length - 8} más...
                                    </div>
                                  )}
                                </div>
                              </div>
                            </div>
                          )
                        })}
                      </div>
                      
                      {stats.length === 0 && (
                        <div className="no-stats">
                          <p>No hay datos suficientes para generar estadísticas</p>
                        </div>
                      )}
                    </>
                  )
                })()
              ) : (
                // Vista normal (no admin o sin datos de hojas)
                <>
                  <div className="stats-summary">
                    <div className="summary-card">
                      <span className="summary-number">{sheetData?.data?.length ?? 0}</span>
                      <span className="summary-label">Total PDV</span>
                    </div>
                    <div className="summary-card summary-green">
                      <span className="summary-number">
                        {sheetData.data.filter(row => {
                          const { relevadorIndex } = getAutoFillIndexes()
                          return relevadorIndex !== -1 && String(row[relevadorIndex] || '').trim() !== ''
                        }).length}
                      </span>
                      <span className="summary-label">Relevados</span>
                    </div>
                    <div className="summary-card summary-purple">
                      <span className="summary-number">
                        {(() => {
                          const { relevadorIndex } = getAutoFillIndexes()
                          const { latIndex, lngIndex } = getLatLngIndexes()
                          if (relevadorIndex === -1 || latIndex === -1 || lngIndex === -1) return 0
                          const parseCoord = (val: unknown): number => {
                            const s = String(val ?? '').trim().replace(',', '.')
                            if (!s) return NaN
                            const n = parseFloat(s)
                            return typeof n === 'number' && !isNaN(n) ? n : NaN
                          }
                          return sheetData.data.filter(row => {
                            if (String(row[relevadorIndex] || '').trim() === '') return false
                            const lat = parseCoord(row[latIndex])
                            const lng = parseCoord(row[lngIndex])
                            return !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0
                          }).length
                        })()}
                      </span>
                      <span className="summary-label">Relevados con Coordenadas</span>
                    </div>
                    <div className="summary-card summary-orange">
                      <span className="summary-number">
                        {sheetData.data.filter(row => {
                          const { relevadorIndex } = getAutoFillIndexes()
                          return relevadorIndex === -1 || String(row[relevadorIndex] || '').trim() === ''
                        }).length}
                      </span>
                      <span className="summary-label">Pendientes</span>
                    </div>
                    <div className="summary-card summary-blue">
                      <span className="summary-number">
                        {sheetData.data.length > 0
                          ? Math.round(
                              (sheetData.data.filter(row => {
                                const { relevadorIndex } = getAutoFillIndexes()
                                return relevadorIndex !== -1 && String(row[relevadorIndex] || '').trim() !== ''
                              }).length / sheetData.data.length) * 100
                            ) + '%'
                          : '0%'}
                      </span>
                      <span className="summary-label">Progreso</span>
                    </div>
                  </div>
                  
                  <div className="charts-grid">
                    {getFieldStats().map((fieldStat, idx) => {
                      const total = fieldStat.data.reduce((sum, d) => sum + d.count, 0)
                      let cumulativePercent = 0
                      
                      const gradientStops = fieldStat.data.map(d => {
                        const percent = (d.count / total) * 100
                        const start = cumulativePercent
                        cumulativePercent += percent
                        return `${d.color} ${start}% ${cumulativePercent}%`
                      }).join(', ')
                      
                      return (
                        <div key={idx} className="chart-card">
                          <h3 className="chart-title">{fieldStat.fieldName}</h3>
                          <div className="chart-content">
                            <div 
                              className="pie-chart"
                              style={{ 
                                background: `conic-gradient(${gradientStops})`
                              }}
                            >
                              <div className="pie-center">
                                <span className="pie-total">{total}</span>
                                <span className="pie-label">Total</span>
                              </div>
                            </div>
                            <div className="chart-legend">
                              {fieldStat.data.slice(0, 8).map((d, i) => (
                                <div key={i} className="legend-item">
                                  <span 
                                    className="legend-color" 
                                    style={{ background: d.color }}
                                  ></span>
                                  <span className="legend-label">{d.label}</span>
                                  <span className="legend-count">{d.count}</span>
                                  <span className="legend-percent">
                                    ({Math.round((d.count / total) * 100)}%)
                                  </span>
                                </div>
                              ))}
                              {fieldStat.data.length > 8 && (
                                <div className="legend-more">
                                  +{fieldStat.data.length - 8} más...
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                  
                  {getFieldStats().length === 0 && (
                    <div className="no-stats">
                      <p>No hay datos suficientes para generar estadísticas</p>
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Edit Permission Modal */}
      {editingPermission && (
        <div className="modal-overlay" onClick={() => {
          setEditingPermission(null)
          setAssignmentMode('ids')
          setSelectedSheet('')
        }}>
          <div className="modal-content modal-medium" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Asignar PDV</h2>
              <button className="modal-close" onClick={() => {
                setEditingPermission(null)
                setAssignmentMode('ids')
                setSelectedSheet('')
              }}>×</button>
            </div>
            <div className="modal-body">
              <p className="modal-user-email">{editingPermission.email}</p>
              
              {/* Mostrar nivel actual (solo lectura - se edita desde la hoja de Sheets) */}
              <div className="user-level-display">
                <span className="level-label">Nivel actual:</span>
                <span className={`user-level-badge level-${editingPermission.level || 1}`}>
                  {editingPermission.level === 3 ? '👑 Admin (Nivel 3)' : editingPermission.level === 2 ? '👁️ Supervisor (Nivel 2)' : '👤 Usuario (Nivel 1)'}
                </span>
                <p className="level-edit-hint">ℹ️ Para cambiar el nivel, editar columna "Nivel" en la hoja Permisos de Google Sheets</p>
              </div>
              
              {/* Selector de modo de asignación */}
              <div className="assignment-mode-selector">
                <label>Modo de asignación de PDV:</label>
                <div className="assignment-mode-buttons">
                  <button
                    type="button"
                    className={`mode-btn ${assignmentMode === 'ids' ? 'active' : ''}`}
                    onClick={() => setAssignmentMode('ids')}
                  >
                    📝 Por IDs
                  </button>
                  <button
                    type="button"
                    className={`mode-btn ${assignmentMode === 'sheet' ? 'active' : ''}`}
                    onClick={() => setAssignmentMode('sheet')}
                  >
                    📋 Por Hoja
                  </button>
                </div>
              </div>

              {assignmentMode === 'ids' ? (
                <div className="edit-field">
                  <label>IDs permitidos (separados por coma)</label>
                  <textarea
                    value={newPermIds}
                    onChange={(e) => setNewPermIds(e.target.value)}
                    placeholder="Ej: 1, 10, 5, 8, 99, 112"
                    rows={4}
                  />
                </div>
              ) : (
                <div className="edit-field">
                  <label>Seleccionar hoja del spreadsheet</label>
                  {loadingSheets ? (
                    <div className="loading-sheets">Cargando hojas...</div>
                  ) : (
                    <>
                      <select
                        value={selectedSheet}
                        onChange={(e) => setSelectedSheet(e.target.value)}
                        className="sheet-select"
                      >
                        <option value="">-- Seleccionar hoja --</option>
                        {availableSheets.map((sheet, idx) => (
                          <option key={idx} value={sheet}>{sheet}</option>
                        ))}
                      </select>
                      {selectedSheet && (
                        <div className="sheet-info">
                          <span className="sheet-info-icon">ℹ️</span>
                          <span>Se asignarán todos los IDs de la hoja "{selectedSheet}" a este usuario.</span>
                        </div>
                      )}
                    </>
                  )}
                </div>
              )}
              
              {/* IDs actuales */}
              {editingPermission.allowedIds.length > 0 && (
                <div className="current-ids-section">
                  <label>IDs actuales asignados ({editingPermission.allowedIds.length}):</label>
                  <div className="current-ids-tags">
                    {editingPermission.allowedIds.slice(0, 20).map((id, i) => (
                      <span key={i} className="id-tag-small">{id}</span>
                    ))}
                    {editingPermission.allowedIds.length > 20 && (
                      <span className="id-tag-more">+{editingPermission.allowedIds.length - 20} más</span>
                    )}
                  </div>
                </div>
              )}
            </div>
            <div className="modal-footer">
              <button 
                className="btn-secondary" 
                onClick={() => {
                  setEditingPermission(null)
                  setAssignmentMode('ids')
                  setSelectedSheet('')
                }} 
                disabled={savingPermissions || loadingSheetIds}
              >
                Cancelar
              </button>
              <button 
                className="btn-primary" 
                onClick={async () => {
                  if (assignmentMode === 'ids') {
                    const ids = newPermIds.split(',').map(id => id.trim()).filter(id => id)
                    handleSavePermissions(editingPermission.email, ids, '') // Sin hoja asignada, nivel se preserva
                  } else if (assignmentMode === 'sheet' && selectedSheet) {
                    // Cargar IDs de la hoja seleccionada y guardar la hoja asignada
                    const sheetIds = await loadSheetIds(selectedSheet)
                    if (sheetIds.length > 0) {
                      handleSavePermissions(editingPermission.email, sheetIds, selectedSheet) // Con hoja asignada, nivel se preserva
                    } else {
                      alert('No se encontraron IDs en la hoja seleccionada')
                      return
                    }
                  } else {
                    alert('Por favor selecciona una hoja')
                    return
                  }
                  setAssignmentMode('ids')
                  setSelectedSheet('')
                }}
                disabled={savingPermissions || loadingSheetIds || (assignmentMode === 'sheet' && !selectedSheet)}
              >
                {savingPermissions || loadingSheetIds ? 'Guardando...' : 'Guardar'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <header className="dashboard-header">
        <div className="header-left">
          <Image 
            src="/Clarín_logo.svg.png" 
            alt="Clarín Logo" 
            width={120} 
            height={35}
            priority
          />
          <h1>Relevamiento de PDV</h1>
        </div>
        <div className="header-right">
          {/* Botones de acción - Solo desktop */}
          <div className="header-action-buttons">
            {/* Selector de hojas para admin */}
            {sheetData?.permissions?.isAdmin && availableSheets.length > 0 && (
              <div className="header-sheet-selector">
                <label className="header-sheet-label">Relevamiento:</label>
                <select
                  value={adminSelectedSheet || availableSheets[0] || ''}
                  onChange={(e) => {
                    setAdminSelectedSheet(e.target.value)
                    setCurrentPage(1)
                    if (accessToken) {
                      loadSheetData(accessToken, e.target.value)
                    }
                  }}
                  className="header-sheet-select"
                  disabled={loadingData}
                >
                  <option value="Todos">📊 Todos</option>
                  {availableSheets
                    .filter(sheet => sheetData?.permissions?.role === 'admin' || sheet.toLowerCase() !== 'test')
                    .map((sheet, idx) => (
                      <option key={idx} value={sheet}>{sheet}</option>
                    ))}
                </select>
              </div>
            )}
            {/* Descargar Reporte, Nuevo PDV y Ver mapa unificados dentro del Panel Admin */}
            {!(sheetData?.permissions?.role === 'admin' || sheetData?.permissions?.role === 'supervisor') && (
              <button 
                className="btn-download-cuestionario"
                onClick={() => setShowNuevoPdvModal(true)}
                title="Agregar nuevo PDV o descargar cuestionario"
              >
                ➕ Nuevo PDV
              </button>
            )}
            {(sheetData?.permissions?.role === 'admin' || sheetData?.permissions?.role === 'supervisor') && (
              <button 
                className="btn-admin"
                onClick={() => {
                  setShowAdminSidebar(true)
                  loadAllSheetsStats() // Cargar stats para el mini resumen
                }}
              >
                ⚙️ Panel Admin
              </button>
            )}
          </div>
          {/* Contador de PDV relevados para usuarios con IDs asignados */}
          {sheetData && userEmail && sheetData.permissions?.allowedIds && sheetData.permissions.allowedIds.length > 0 && !sheetData.permissions.isAdmin && (() => {
            const stats = getUserStats(userEmail, sheetData.permissions.allowedIds)
            const percent = stats.total > 0 ? Math.round((stats.relevados / stats.total) * 100) : 0
            return (
              <div className="pdv-counter">
                <div className="pdv-counter-label">PDV Relevados</div>
                <div className="pdv-counter-value">
                  <span className="pdv-done">{stats.relevados}</span>
                  <span className="pdv-separator">/</span>
                  <span className="pdv-total">{stats.total}</span>
                </div>
                <div className="pdv-counter-bar">
                  <div className="pdv-counter-fill" style={{ width: `${percent}%` }}></div>
                </div>
              </div>
            )
          })()}
          <div className="user-badge">
            <span className="user-email">{userEmail}</span>
            {sheetData?.permissions?.role === 'admin' && (
              <span className="admin-badge">Admin</span>
            )}
            {sheetData?.permissions?.role === 'supervisor' && (
              <span className="admin-badge supervisor-badge">Supervisor</span>
            )}
          </div>
          <button className="btn-secondary" onClick={handleSignoutClick}>
            Cerrar Sesión
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="dashboard-main">
        {/* Mobile Location Filters Bar - Solo visible en móvil */}
        <div className="mobile-top-filters">
          {partidoOptions.length > 0 && (
            <>
              <select 
                value={filterPartido}
                onChange={(e) => setFilterPartido(e.target.value)}
                className="mobile-top-select"
              >
                <option value="">📍 Todos</option>
                {partidoOptions.map((partido, idx) => (
                  <option key={idx} value={partido}>{partido}</option>
                ))}
              </select>
              
              <select 
                value={filterLocalidad}
                onChange={(e) => setFilterLocalidad(e.target.value)}
                className="mobile-top-select"
                disabled={!filterPartido}
              >
                <option value="">🏘️ Localidad</option>
                {localidadOptions.map((loc, idx) => (
                  <option key={idx} value={loc}>{loc}</option>
                ))}
              </select>
              
              {(filterPartido || filterLocalidad) && (
                <button 
                  className="mobile-top-clear"
                  onClick={() => {
                    setFilterPartido('')
                    setFilterLocalidad('')
                  }}
                >
                  ✕
                </button>
              )}
            </>
          )}
          
          <button 
            className={`mobile-top-reload ${loadingData ? 'loading' : ''}`}
            onClick={() => {
              if (accessToken && !loadingData) {
                loadSheetData(accessToken, adminSelectedSheet || '')
              }
            }}
            disabled={loadingData}
            title="Recargar datos"
          >
            🔄
          </button>
        </div>

        {/* Toolbar */}
        <div className="toolbar">
          <div className="toolbar-left">
            <div className="search-group">
              <select 
                value={searchType}
                onChange={(e) => setSearchType(e.target.value as 'id' | 'paquete' | 'direccion')}
                className="search-select"
              >
                <option value="id">ID</option>
                <option value="paquete">Paquete</option>
                <option value="direccion">Dirección</option>
              </select>
              <input
                type="text"
                placeholder={
                  searchType === 'id' ? 'Buscar por ID exacto...' :
                  searchType === 'paquete' ? 'Buscar paquetes que contengan...' :
                  'Buscar por dirección...'
                }
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="search-input"
              />
              {searchTerm && (
                <button 
                  className="search-clear"
                  onClick={() => setSearchTerm('')}
                  title="Limpiar búsqueda"
                >
                  ×
                </button>
              )}
            </div>
            <select 
              value={filterRelevado}
              onChange={(e) => setFilterRelevado(e.target.value as 'todos' | 'relevados' | 'no_relevados' | 'censados_sin_mapear' | 'censados_mapeados')}
              className="filter-relevado-select"
            >
              <option value="todos">📋 Todos los PDV</option>
              <option value="relevados">✅ Solo relevados</option>
              <option value="no_relevados">⏳ Sin relevar</option>
              <option value="censados_mapeados">📍 Censados Mapeados</option>
              <option value="censados_sin_mapear">📍 Censados sin Mapear</option>
            </select>
            
            {/* Filtros por ubicación - Desktop */}
            {partidoOptions.length > 0 && (
              <div className="desktop-location-filters">
                <select 
                  value={filterPartido}
                  onChange={(e) => setFilterPartido(e.target.value)}
                  className="filter-location-select"
                >
                  <option value="">📍 Todos</option>
                  {partidoOptions.map((partido, idx) => (
                    <option key={idx} value={partido}>{partido}</option>
                  ))}
                </select>
                
                <select 
                  value={filterLocalidad}
                  onChange={(e) => setFilterLocalidad(e.target.value)}
                  className="filter-location-select"
                  disabled={!filterPartido || localidadOptions.length === 0}
                >
                  <option value="">🏘️ Localidad</option>
                  {localidadOptions.map((loc, idx) => (
                    <option key={idx} value={loc}>{loc}</option>
                  ))}
                </select>
                
                {(filterPartido || filterLocalidad) && (
                  <button 
                    className="btn-clear-location-filters"
                    onClick={() => {
                      setFilterPartido('')
                      setFilterLocalidad('')
                    }}
                    title="Limpiar filtros de ubicación"
                  >
                    ✕
                  </button>
                )}
              </div>
            )}
            
            {/* Indicador de hoja asignada para usuarios comunes (sin acceso a ALTA PDV) */}
            {!sheetData?.permissions?.isAdmin && sheetData?.permissions?.assignedSheet && (
              <div className="sheet-filter-group user-sheet-indicator">
                <span className="sheet-filter-label">📋 Hoja:</span>
                <span className="sheet-assigned-name">{sheetData.permissions.assignedSheet}</span>
              </div>
            )}
            <button 
              className="btn-primary"
              onClick={() => accessToken && loadSheetData(accessToken, adminSelectedSheet)}
              disabled={loadingData}
            >
              {loadingData ? 'Cargando...' : 'Recargar Datos'}
            </button>
          </div>
          <div className="toolbar-right">
            {sheetData && sortedData.length > 0 && (
              <span className="results-count">
                Mostrando {startIndex + 1} a {Math.min(endIndex, sortedData.length)} de {sortedData.length} registros
              </span>
            )}
          </div>
        </div>

        {/* Error Message */}
        {error && (
          <div className="error-message">
            <p>
              {error.includes('Error interno del servidor')
                ? 'Error interno del servidor. Cierre la web y vuelva a ingresar.'
                : error}
            </p>
            <button
              type="button"
              className="btn-secondary error-signout-btn"
              onClick={() => {
                setError(null)
                handleSignoutClick()
              }}
            >
              Cerrar sesión
            </button>
          </div>
        )}

        {/* Loading State */}
        {loadingData && (
          <div className="loading-container">
            <div className="spinner-large"></div>
            <p>Cargando datos...</p>
            {currentTip && (
              <div className="loading-tip">
                <span className="tip-icon">💡</span>
                <span className="tip-text">
                  <strong>¿Sabías que?</strong> {currentTip}
                </span>
              </div>
            )}
          </div>
        )}

        {/* Data Table */}
        {!loadingData && sheetData && (
          <div className="table-container">
            {sortedData.length === 0 ? (
              <div className="no-data-container">
                <p>No hay datos disponibles</p>
                {!sheetData.permissions?.isAdmin && (
                  <p className="hint">Es posible que no tengas permisos asignados para ver registros.</p>
                )}
              </div>
            ) : (
              <table className="data-table">
                <thead>
                  <tr>
                    <th className="actions-col">Acciones</th>
                    {sheetData.headers.map((header, idx) => {
                      // Ocultar columna DISPOSITIVO en la tabla
                      const headerLower = header.toLowerCase().trim()
                      if (headerLower === 'dispositivo' || headerLower.includes('dispositivo')) return null
                      // Ocultar columna __HOJA__ para usuarios no admin, y renombrarla para admins
                      if (header === '__HOJA__') {
                        if (!sheetData.permissions.isAdmin) return null
                        return (
                          <th 
                            key={idx} 
                            className={`sortable-header hoja-column ${sortColumn === idx ? 'sorted' : ''}`}
                            onClick={() => handleColumnSort(idx)}
                            title="Ordenar por Hoja"
                          >
                            <span className="header-content">
                              📋 HOJA
                              <span className="sort-indicator">
                                {sortColumn === idx ? (
                                  sortDirection === 'asc' ? ' ▲' : ' ▼'
                                ) : ' ⇅'}
                              </span>
                            </span>
                          </th>
                        )
                      }
                      return (
                        <th 
                          key={idx} 
                          className={`sortable-header ${sortColumn === idx ? 'sorted' : ''}`}
                          onClick={() => handleColumnSort(idx)}
                          title={`Ordenar por ${header}`}
                        >
                          <span className="header-content">
                            {header}
                            <span className="sort-indicator">
                              {sortColumn === idx ? (
                                sortDirection === 'asc' ? ' ▲' : ' ▼'
                              ) : ' ⇅'}
                            </span>
                          </span>
                        </th>
                      )
                    })}
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.map((row, rowIdx) => {
                    const originalIndex = sheetData.data.indexOf(row)
                    
                    // Verificar si la fila ya fue relevada (tiene valor en "Relevador por:")
                    const { relevadorIndex } = getAutoFillIndexes()
                    const relevadorValue = relevadorIndex !== -1 ? String(row[relevadorIndex] || '').trim() : ''
                    const isRelevado = relevadorValue !== ''
                    
                    return (
                      <tr key={rowIdx} className={isRelevado ? 'row-relevado' : ''}>
                        <td className="actions-col">
                          <div className="actions-buttons">
                            <button 
                              className="btn-edit"
                              onClick={() => handleEditRow(originalIndex)}
                              title="Editar registro"
                            >
                              ✎
                            </button>
                            <button 
                              className="btn-location"
                              onClick={() => handleLocationClick(originalIndex)}
                              disabled={savingLocationRow === originalIndex}
                              title="Opciones de ubicación"
                            >
                              {savingLocationRow === originalIndex ? '⏳' : '📍'}
                            </button>
                          </div>
                        </td>
                        {row.map((cell, cellIdx) => {
                          // Ocultar columna DISPOSITIVO en la tabla
                          const headerLower = sheetData.headers[cellIdx]?.toLowerCase().trim() || ''
                          if (headerLower === 'dispositivo' || headerLower.includes('dispositivo')) return null
                          // Ocultar columna __HOJA__ para usuarios no admin
                          if (sheetData.headers[cellIdx] === '__HOJA__') {
                            if (!sheetData.permissions.isAdmin) return null
                            return <td key={cellIdx} className="hoja-cell">{String(cell || '')}</td>
                          }
                          return <td key={cellIdx}>{String(cell || '')}</td>
                        })}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* Pagination */}
        {!loadingData && sheetData && sortedData.length > 0 && (
          <div className="pagination-container">
            <div className="pagination-info">
              Mostrando {startIndex + 1} a {Math.min(endIndex, sortedData.length)} de {sortedData.length} registros
            </div>
            <div className="pagination-controls">
              <button 
                className="pagination-btn"
                onClick={() => setCurrentPage(1)}
                disabled={currentPage === 1}
              >
                Primera
              </button>
              <button 
                className="pagination-btn"
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
              >
                Anterior
              </button>
              
              {getPageNumbers().map((page, idx) => (
                typeof page === 'number' ? (
                  <button
                    key={idx}
                    className={`pagination-btn pagination-num ${currentPage === page ? 'active' : ''}`}
                    onClick={() => setCurrentPage(page)}
                  >
                    {page}
                  </button>
                ) : (
                  <span key={idx} className="pagination-ellipsis">...</span>
                )
              ))}
              
      <button
                className="pagination-btn"
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
              >
                Siguiente
      </button>
      <button
                className="pagination-btn"
                onClick={() => setCurrentPage(totalPages)}
                disabled={currentPage === totalPages}
              >
                Última
      </button>
            </div>
          </div>
        )}
      </main>

      {/* Mobile Bottom Navigation */}
      {isAuthorized && (
        <nav className="mobile-bottom-nav">
          <button 
            className="mobile-nav-item"
            onClick={() => {
              // Si hay búsqueda con texto, preguntar confirmación
              if (showMobileSearch && mobileSearchQuery.trim()) {
                const confirmed = window.confirm('¿Deseas salir de la búsqueda? Se perderá el texto ingresado.')
                if (!confirmed) return
              }
              setShowMobileSearch(false)
              setShowMobileStats(false)
              setShowAdminSidebar(false)
              setShowNuevoPdvModal(true)
            }}
          >
            <span className="nav-icon">⚠️</span>
            <span className="nav-label">Ayuda</span>
          </button>
          
          <button 
            className={`mobile-nav-item ${showMobileSearch ? 'active' : ''}`}
            onClick={() => {
              setShowMobileSearch(true)
              setMobileSearchQuery('')
              setShowMobileStats(false)
              setShowAdminSidebar(false)
              setShowNuevoPdvModal(false)
            }}
          >
            <span className="nav-icon">🔍</span>
            <span className="nav-label">Buscar</span>
          </button>
          
          <button 
            className={`mobile-nav-item nav-center ${!showAdminSidebar && !editingRow && locationModalRow === null ? 'active' : ''}`}
            onClick={() => {
              // Si hay edición, ubicación abierta, o formulario de nuevo PDV, pedir confirmación
              if (editingRow !== null || locationModalRow !== null || showMapPicker || showNuevoPdvForm) {
                const confirmed = window.confirm(
                  '⚠️ Tienes cambios sin guardar.\n\n¿Deseas volver al inicio? Los cambios no guardados se perderán.'
                )
                if (!confirmed) return
              }
              // Si hay búsqueda con texto, preguntar confirmación
              if (showMobileSearch && mobileSearchQuery.trim()) {
                const confirmed = window.confirm('¿Deseas salir de la búsqueda? Se perderá el texto ingresado.')
                if (!confirmed) return
              }
              
              setShowAdminSidebar(false)
              setEditingRow(null)
              setEditedValues([])
              setLocationModalRow(null)
              setShowMapPicker(false)
              setManualCoords(null)
              setShowLocationWarning(false)
              setShowMobileSearch(false)
              setMobileSearchQuery('')
              setShowMobileStats(false)
              setShowNuevoPdvModal(false)
              setShowNuevoPdvForm(false)
              resetNuevoPdvForm()
              // Recargar datos
              if (accessToken && !loadingData) {
                loadSheetData(accessToken, adminSelectedSheet || '')
              }
            }}
          >
            <Image 
              src="/Clarinpng.png" 
              alt="Inicio" 
              width={55} 
              height={55}
              className="nav-center-logo"
            />
          </button>
          
          {sheetData?.permissions?.isAdmin ? (
            <button 
              className={`mobile-nav-item ${showAdminSidebar ? 'active' : ''}`}
              onClick={() => {
                // Si hay búsqueda con texto, preguntar confirmación
                if (showMobileSearch && mobileSearchQuery.trim()) {
                  const confirmed = window.confirm('¿Deseas salir de la búsqueda? Se perderá el texto ingresado.')
                  if (!confirmed) return
                }
                setShowMobileSearch(false)
                setShowMobileStats(false)
                setShowNuevoPdvModal(false)
                setShowAdminSidebar(!showAdminSidebar)
                if (!showAdminSidebar) loadAllSheetsStats()
              }}
            >
              <span className="nav-icon">⚙️</span>
              <span className="nav-label">Admin</span>
            </button>
          ) : (
            <button 
              className={`mobile-nav-item ${showMobileStats ? 'active' : ''}`}
              onClick={() => {
                // Si hay búsqueda con texto, preguntar confirmación
                if (showMobileSearch && mobileSearchQuery.trim()) {
                  const confirmed = window.confirm('¿Deseas salir de la búsqueda? Se perderá el texto ingresado.')
                  if (!confirmed) return
                }
                setShowMobileSearch(false)
                setShowAdminSidebar(false)
                setShowNuevoPdvModal(false)
                setShowMobileStats(!showMobileStats)
              }}
            >
              <span className="nav-icon">📊</span>
              <span className="nav-label">Stats</span>
            </button>
          )}
          
          <button 
            className="mobile-nav-item"
            onClick={() => {
              if (window.confirm('¿Deseas cerrar sesión?')) {
                handleSignoutClick()
              }
            }}
          >
            <span className="nav-icon">🚪</span>
            <span className="nav-label">Salir</span>
          </button>
        </nav>
      )}

      {/* Mobile Stats Modal */}
      {showMobileStats && sheetData && (
        <div className="mobile-stats-modal">
          <button className="stats-close" onClick={() => setShowMobileStats(false)}>✕</button>
          <div className="stats-header">
            <h3>📊 Tu Progreso</h3>
          </div>
          {(() => {
            const allowedIds = sheetData.permissions?.allowedIds || []
            const total = allowedIds.length > 0 ? allowedIds.length : (sheetData?.data?.length ?? 0)
            const { relevadorIndex } = getAutoFillIndexes()
            let relevados = 0
            if (relevadorIndex !== -1) {
              relevados = sheetData.data.filter(row => String(row[relevadorIndex] || '').trim() !== '').length
            }
            const faltantes = total - relevados
            const progreso = total > 0 ? Math.round((relevados / total) * 100) : 0
            
            return (
              <>
                <div className="stats-grid">
                  <div className="stat-item relevados">
                    <span className="stat-number">{relevados}</span>
                    <span className="stat-label">Relevados</span>
                  </div>
                  <div className="stat-item pendientes">
                    <span className="stat-number">{faltantes}</span>
                    <span className="stat-label">Pendientes</span>
                  </div>
                  <div className="stat-item total">
                    <span className="stat-number">{total}</span>
                    <span className="stat-label">Total</span>
                  </div>
                </div>
                <div className="progress-section">
                  <div className="progress-label">Progreso General</div>
                  <div className="progress-value">{progreso}%</div>
                </div>
              </>
            )
          })()}
        </div>
      )}

      {/* Mobile Search Modal */}
      {showMobileSearch && sheetData && (
        <div className="mobile-search-modal">
          <div className="search-header">
            <button className="search-back" onClick={() => setShowMobileSearch(false)}>←</button>
            <div className="search-input-wrapper">
              <span className="search-icon">🔍</span>
              <input
                type="text"
                placeholder="Buscar..."
                value={mobileSearchQuery}
                onChange={(e) => setMobileSearchQuery(e.target.value)}
                autoFocus
              />
            </div>
          </div>
          
          <div className="search-filters">
            <button 
              className={`filter-chip ${mobileSearchType === 'id' ? 'active' : ''}`}
              onClick={() => setMobileSearchType('id')}
            >
              Por ID
            </button>
            <button 
              className={`filter-chip ${mobileSearchType === 'paquete' ? 'active' : ''}`}
              onClick={() => setMobileSearchType('paquete')}
            >
              Por Paquete
            </button>
            <button 
              className={`filter-chip ${mobileSearchType === 'direccion' ? 'active' : ''}`}
              onClick={() => setMobileSearchType('direccion')}
            >
              Por Dirección
            </button>
          </div>
          
          <div className="search-results">
            {(() => {
              const query = mobileSearchQuery.toLowerCase().trim()
              if (!query) {
                return (
                  <div className="no-results">
                    <div className="no-results-icon">🔍</div>
                    <p>Escribe para buscar registros</p>
                  </div>
                )
              }
              
              const paqueteIndex = sheetData.headers.findIndex(h => 
                h.toLowerCase().includes('paquete')
              )
              const { relevadorIndex } = getAutoFillIndexes()
              
              const partidoIdx = getPartidoIndex()
              const localidadIdx = getLocalidadIndex()
              
              const domicilioIndex = sheetData.headers.findIndex((h: string) =>
                h.toLowerCase().includes('domicilio') || h.toLowerCase() === 'direccion' || h.toLowerCase() === 'dirección'
              )
              const results = sheetData.data.filter((row, idx) => {
                // Filtro por texto
                if (mobileSearchType === 'id') {
                  if (!String(row[0] || '').toLowerCase().includes(query)) return false
                } else if (mobileSearchType === 'paquete') {
                  if (paqueteIndex === -1 || !String(row[paqueteIndex] || '').toLowerCase().includes(query)) return false
                } else if (mobileSearchType === 'direccion') {
                  if (domicilioIndex === -1 || !String(row[domicilioIndex] || '').toLowerCase().includes(query)) return false
                }
                
                // Filtro por Partido (case-insensitive)
                if (filterPartido && partidoIdx !== -1) {
                  if (String(row[partidoIdx] || '').trim().toLowerCase() !== filterPartido.toLowerCase()) return false
                }
                
                // Filtro por Localidad (case-insensitive)
                if (filterLocalidad && localidadIdx !== -1) {
                  if (String(row[localidadIdx] || '').trim().toLowerCase() !== filterLocalidad.toLowerCase()) return false
                }
                
                return true
              }).slice(0, 20) // Limitar a 20 resultados
              
              if (results.length === 0) {
                return (
                  <div className="no-results">
                    <div className="no-results-icon">😕</div>
                    <p>No se encontraron resultados</p>
                  </div>
                )
              }
              
              return results.map((row, idx) => {
                const originalIndex = sheetData.data.indexOf(row)
                const isRelevado = relevadorIndex !== -1 && String(row[relevadorIndex] || '').trim() !== ''
                const paquete = paqueteIndex !== -1 ? String(row[paqueteIndex] || '') : ''
                
                return (
                  <div 
                    key={idx}
                    className="search-result-item"
                  >
                    <div className="result-info">
                      <div className="result-id">ID: {row[0]}</div>
                      {paquete && <div className="result-details">Paquete: {paquete}</div>}
                    </div>
                    <span className={`result-status ${isRelevado ? 'relevado' : 'pendiente'}`}>
                      {isRelevado ? '✓ Relevado' : 'Pendiente'}
                    </span>
                    <div className="result-actions">
                      <button 
                        className="result-action-btn edit"
                        onClick={() => {
                          setShowMobileSearch(false)
                          setCameFromMobileSearch(true)
                          handleEditRow(originalIndex)
                        }}
                        title="Editar"
                      >
                        ✎
                      </button>
                      <button 
                        className="result-action-btn location"
                        onClick={() => {
                          setShowMobileSearch(false)
                          setCameFromMobileSearch(true)
                          handleLocationClick(originalIndex)
                        }}
                        title="Ubicación"
                      >
                        📍
                      </button>
                    </div>
                  </div>
                )
              })
            })()}
          </div>
        </div>
      )}
    </div>
  )
}

