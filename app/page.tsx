'use client'

import { useEffect, useRef, useState } from 'react'

declare global {
  interface Window {
    gapi: any
    google: any
    gapiLoaded: () => void
    gisLoaded: () => void
    handleAuthClick: () => void
    handleSignoutClick: () => void
  }
}

export default function Home() {
  const [data, setData] = useState<any[][]>([])
  const [headers, setHeaders] = useState<string[]>([])
  const [searchTerm, setSearchTerm] = useState('')
  const [searchType, setSearchType] = useState<'id' | 'paquete'>('id')
  const [currentPage, setCurrentPage] = useState(1)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [isAuthenticated, setIsAuthenticated] = useState(false)
  const [editingRowId, setEditingRowId] = useState<string | null>(null)
  const [editFormData, setEditFormData] = useState<any[]>([])
  const [saving, setSaving] = useState(false)
  const [userEmail, setUserEmail] = useState<string>('')
  const [isAdmin, setIsAdmin] = useState(false)
  const [showAdminPanel, setShowAdminPanel] = useState(false)
  const [adminStats, setAdminStats] = useState<{ email: string; relevados: number; faltantes: number }[]>([])
  const [userPermissions, setUserPermissions] = useState<{ allowedIds: string[]; isAdmin: boolean }>({ allowedIds: [], isAdmin: false })
  const [allPermissions, setAllPermissions] = useState<Array<{ email: string; allowedIds: string[] }>>([])
  const [showPermissionsModal, setShowPermissionsModal] = useState(false)
  const [selectedUserForPermission, setSelectedUserForPermission] = useState<string>('')
  const [permissionIdsInput, setPermissionIdsInput] = useState<string>('')
  const authorizeButtonRef = useRef<HTMLButtonElement>(null)
  const signoutButtonRef = useRef<HTMLButtonElement>(null)
  
  const ITEMS_PER_PAGE = 50
  const SPREADSHEET_ID = '13Ht_fOQuLHDMNYqKFr3FjedtU9ZkKOp_2_zCOnjHKm8'
  
  // Lista de emails autorizados para acceso admin
  const ADMIN_EMAILS = [
    'atoledo.agea@gmail.com'
  ]

  const CLIENT_ID = '549677208908-9h6q933go4ss870pbq8gd8gaae75k338.apps.googleusercontent.com'
  const API_KEY = 'AIzaSyCJUD23abF8LcZPp7e8eiK0D5IfFoRCxUc'
  const DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
  const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email'

  let tokenClient: any
  let gapiInited = false
  let gisInited = false

  useEffect(() => {
    // Initialize buttons visibility
    // Verificar si ya hay un token existente
    const existingToken = window.gapi?.client?.getToken()
    if (existingToken) {
      console.log('Token existente encontrado, autenticando automáticamente...')
      setIsAuthenticated(true)
      // Cargar datos si ya está autenticado
      loadData()
    }

    // Define global functions
    window.gapiLoaded = gapiLoaded
    window.gisLoaded = gisLoaded
    window.handleAuthClick = handleAuthClick
    window.handleSignoutClick = handleSignoutClick

    // Load Google APIs
    const gapiScript = document.createElement('script')
    gapiScript.src = 'https://apis.google.com/js/api.js'
    gapiScript.async = true
    gapiScript.defer = true
    gapiScript.onload = () => window.gapiLoaded()
    document.head.appendChild(gapiScript)

    const gisScript = document.createElement('script')
    gisScript.src = 'https://accounts.google.com/gsi/client'
    gisScript.async = true
    gisScript.defer = true
    gisScript.onload = () => window.gisLoaded()
    document.head.appendChild(gisScript)

    return () => {
      // Cleanup if needed
    }
  }, [])

  function gapiLoaded() {
    if (window.gapi) {
      window.gapi.load('client', initializeGapiClient)
    }
  }

  async function initializeGapiClient() {
    await window.gapi.client.init({
      apiKey: API_KEY,
      discoveryDocs: [DISCOVERY_DOC],
    })
    gapiInited = true
    maybeEnableButtons()
  }

  function gisLoaded() {
    if (window.google) {
      tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: '', // defined later
      })
      gisInited = true
      maybeEnableButtons()
    }
  }

  function maybeEnableButtons() {
    if (gapiInited && gisInited && authorizeButtonRef.current) {
      authorizeButtonRef.current.style.visibility = 'visible'
    }
  }

  async function handleAuthClick() {
    console.log('handleAuthClick ejecutado - Iniciando autenticación')

    if (!tokenClient) {
      console.log('No hay tokenClient')
      return
    }

    tokenClient.callback = async (resp: any) => {
      if (resp.error !== undefined) {
        setError(`Error de autenticación: ${resp.error}`)
        return
      }
      setIsAuthenticated(true)
      
      // Cargar permisos después de autenticarse
      setTimeout(() => {
        loadPermissions()
      }, 1000)
      
      // Obtener el email del usuario autenticado
      try {
        // Cargar la API de oauth2 si no está cargada
        await window.gapi.client.load('oauth2', 'v2')
        const userInfo = await window.gapi.client.oauth2.userinfo.get()
        if (userInfo.result && userInfo.result.email) {
          const email = userInfo.result.email
          setUserEmail(email)
          console.log('Email obtenido:', email)
          // Verificar si el usuario es admin
          const adminStatus = ADMIN_EMAILS.includes(email.toLowerCase())
          setIsAdmin(adminStatus)
          console.log('Usuario es admin:', adminStatus)
        } else {
          console.log('No se pudo obtener el email del usuario')
        }
      } catch (err) {
        console.error('Error al obtener email del usuario:', err)
        // Intentar método alternativo
        try {
          const token = window.gapi.client.getToken()
          if (token && token.access_token) {
            const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
              headers: {
                'Authorization': `Bearer ${token.access_token}`
              }
            })
            const userData = await response.json()
            if (userData.email) {
              const email = userData.email
              setUserEmail(email)
              console.log('Email obtenido (método alternativo):', email)
              // Verificar si el usuario es admin
              const adminStatus = ADMIN_EMAILS.includes(email.toLowerCase())
              setIsAdmin(adminStatus)
              console.log('Usuario es admin:', adminStatus)
            }
          }
        } catch (altErr) {
          console.error('Error en método alternativo:', altErr)
        }
      }
      
      if (signoutButtonRef.current) {
        signoutButtonRef.current.style.visibility = 'visible'
      }
      if (authorizeButtonRef.current) {
        authorizeButtonRef.current.innerText = 'Actualizar Datos'
      }
      // loadData se llamará después de cargar permisos
      await loadData()
    }

    if (window.gapi.client.getToken() === null) {
      tokenClient.requestAccessToken({ prompt: 'consent' })
    } else {
      tokenClient.requestAccessToken({ prompt: '' })
    }
  }

  function handleSignoutClick() {
    const token = window.gapi.client.getToken()
    if (token !== null) {
      window.google.accounts.oauth2.revoke(token.access_token)
      window.gapi.client.setToken('')
      setData([])
      setHeaders([])
      setError(null)
      setIsAuthenticated(false)
      setUserEmail('')
      setIsAdmin(false)
      if (authorizeButtonRef.current) {
        authorizeButtonRef.current.innerText = 'Iniciar Sesión'
      }
      if (signoutButtonRef.current) {
        signoutButtonRef.current.style.visibility = 'hidden'
      }
    }
  }

  // Función helper para obtener el token de acceso
  const getAccessToken = () => {
    const token = window.gapi.client.getToken()
    return token?.access_token || null
  }

  // Cargar permisos del usuario
  async function loadPermissions() {
    const token = getAccessToken()
    if (!token) return

    try {
      const response = await fetch('/api/permissions', {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      })

      if (!response.ok) {
        throw new Error('Error al cargar permisos')
      }

      const data = await response.json()
      
      if (data.isAdmin) {
        setUserPermissions({ allowedIds: [], isAdmin: true })
        setAllPermissions(data.permissions || [])
      } else {
        setUserPermissions({ allowedIds: data.allowedIds || [], isAdmin: false })
      }
    } catch (err: any) {
      console.error('Error cargando permisos:', err)
    }
  }

  async function loadData() {
    setLoading(true)
    setError(null)
    
    const token = getAccessToken()
    if (!token) {
      setError('No estás autenticado')
      setLoading(false)
      return
    }

    try {
      const response = await fetch('/api/data', {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error al cargar datos')
      }

      const result = await response.json()

      // Guardar permisos
      setUserPermissions(result.permissions || { allowedIds: [], isAdmin: false })
      setIsAdmin(result.permissions?.isAdmin || false)

      // Procesar datos
      if (result.headers && result.headers.length > 0) {
        setHeaders(result.headers)
        setData(result.data || [])
      } else {
        setHeaders([])
        setData([])
      }

      // Si es admin, también cargar todos los permisos
      if (result.permissions?.isAdmin) {
        await loadPermissions()
      }
    } catch (err: any) {
      const errorMessage = err.message || 'Error desconocido al cargar datos'
      setError(`Error: ${errorMessage}`)
      console.error('Error cargando datos:', err)
    } finally {
      setLoading(false)
    }
  }

  // Filtrar datos según el tipo de búsqueda seleccionado
  const filteredData = searchTerm.trim() === ''
    ? data
    : data.filter((row) => {
        if (!row || row.length === 0) return false

        const searchValue = searchTerm.trim().toLowerCase()

        if (searchType === 'id') {
          // Buscar en la primera columna (ID) - comparación exacta
          const idCell = row[0]
          if (idCell === null || idCell === undefined) return false
          return String(idCell).trim().toLowerCase() === searchValue
        } else if (searchType === 'paquete') {
          // Buscar en la columna que contiene "N Paquete" o "Paquete" - comparación exacta
          // Vamos a buscar en todas las columnas que puedan contener información de paquete
          return row.some((cell, index) => {
            if (cell === null || cell === undefined) return false
            const cellValue = String(cell).trim().toLowerCase()
            const header = headers[index]?.toLowerCase() || ''

            // Buscar en columnas que contengan "paquete" en el header
            if (header.includes('paquete') || header.includes('n°') || header.includes('numero')) {
              return cellValue === searchValue
            }
            // También buscar en celdas que parezcan contener números de paquete (comparación exacta)
            if (/^\d+$/.test(cellValue) && searchValue.match(/^\d+$/)) {
              return cellValue === searchValue
            }
            return false
          })
        }

        return false
      })

  // Calcular paginación
  const totalPages = Math.ceil(filteredData.length / ITEMS_PER_PAGE)
  const startIndex = (currentPage - 1) * ITEMS_PER_PAGE
  const endIndex = startIndex + ITEMS_PER_PAGE
  const paginatedData = filteredData.slice(startIndex, endIndex)

  // Resetear a página 1 cuando cambia la búsqueda
  useEffect(() => {
    setCurrentPage(1)
  }, [searchTerm, searchType])

  // Generar números de página para mostrar
  const getPageNumbers = () => {
    const pages: (number | string)[] = []
    const maxVisible = 5
    
    if (totalPages <= maxVisible) {
      // Mostrar todas las páginas si son pocas
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i)
      }
    } else {
      // Mostrar páginas con elipsis
      if (currentPage <= 3) {
        // Al inicio
        for (let i = 1; i <= 4; i++) {
          pages.push(i)
        }
        pages.push('...')
        pages.push(totalPages)
      } else if (currentPage >= totalPages - 2) {
        // Al final
        pages.push(1)
        pages.push('...')
        for (let i = totalPages - 3; i <= totalPages; i++) {
          pages.push(i)
        }
      } else {
        // En el medio
        pages.push(1)
        pages.push('...')
        for (let i = currentPage - 1; i <= currentPage + 1; i++) {
          pages.push(i)
        }
        pages.push('...')
        pages.push(totalPages)
      }
    }
    return pages
  }

  // Función para obtener la fecha actual en formato DD/MM/YYYY
  const getCurrentDate = () => {
    const today = new Date()
    const day = String(today.getDate()).padStart(2, '0')
    const month = String(today.getMonth() + 1).padStart(2, '0')
    const year = today.getFullYear()
    return `${day}/${month}/${year}`
  }

  // Encontrar el índice de la columna "fecha relevada"
  const fechaRelevadaIndex = headers.findIndex((header) => 
    header && header.toLowerCase().includes('fecha relevada')
  )

  // Encontrar el índice de la columna "relevado por"
  const relevadoPorIndex = headers.findIndex((header) => {
    if (!header) return false
    const headerLower = header.toLowerCase().trim()
    return headerLower.includes('relevado por') || headerLower.includes('relevado por') || headerLower === 'relevado por'
  })
  
  // Debug: mostrar el índice encontrado
  if (relevadoPorIndex !== -1) {
    console.log('Columna "Relevado por" encontrada en índice:', relevadoPorIndex, 'Nombre:', headers[relevadoPorIndex])
  } else {
    console.warn('Columna "Relevado por" no encontrada. Headers disponibles:', headers)
  }

  // Función para calcular estadísticas de admin
  const calculateAdminStats = () => {
    if (relevadoPorIndex === -1) {
      console.warn('Columna "Relevado por" no encontrada')
      return
    }

    const stats: { [key: string]: number } = {}
    const emailMap: { [key: string]: string } = {} // Para mantener el email original con mayúsculas

    // Contar registros relevados por cada usuario
    // Solo cuenta si el email está presente y no está vacío en la columna "Relevado por"
    // Y además verifica que el ID del registro esté en los IDs asignados del usuario
    data.forEach(row => {
      // Verificar que la fila tenga al menos la primera columna (ID)
      if (!row || row.length === 0 || row[0] === null || row[0] === undefined) {
        return
      }

      const rowId = String(row[0]).trim().toLowerCase()

      // Verificar que la columna "Relevado por" existe y tiene un valor
      if (row && row.length > relevadoPorIndex && row[relevadoPorIndex] !== null && row[relevadoPorIndex] !== undefined) {
        const emailOriginal = String(row[relevadoPorIndex]).trim()
        
        // Solo contar si el email no está vacío (después de trim)
        if (emailOriginal && emailOriginal.length > 0) {
          const emailLower = emailOriginal.toLowerCase()
          
          // Obtener los IDs asignados a este usuario
          const userPerms = allPermissions.find(p => p.email.toLowerCase() === emailLower)
          if (userPerms && userPerms.allowedIds.length > 0) {
            // Convertir IDs asignados a minúsculas para comparación
            const allowedIdsLower = userPerms.allowedIds.map(id => String(id).trim().toLowerCase())
            
            // Solo contar si el ID de esta fila está en los IDs asignados del usuario
            if (allowedIdsLower.includes(rowId)) {
              stats[emailLower] = (stats[emailLower] || 0) + 1
              // Guardar el email original (con mayúsculas) para mostrarlo
              if (!emailMap[emailLower]) {
                emailMap[emailLower] = emailOriginal
              }
            }
          }
        }
      }
    })

    // Crear array de estadísticas usando los permisos asignados
    const statsArray = allPermissions.map(perm => {
      const emailLower = perm.email.toLowerCase()
      const relevados = stats[emailLower] || 0
      const idsAsignados = perm.allowedIds.length
      
      // Faltantes = IDs asignados - IDs ya relevados
      const faltantes = idsAsignados - relevados
      
      return {
        email: emailMap[emailLower] || perm.email, // Usar el email original si existe, sino el del permiso
        relevados: relevados,
        faltantes: faltantes
      }
    })

    // También incluir usuarios que tienen registros relevados pero no tienen permisos asignados
    Object.keys(stats).forEach(emailLower => {
      const existsInStats = statsArray.find(s => s.email.toLowerCase() === emailLower)
      if (!existsInStats) {
        statsArray.push({
          email: emailMap[emailLower] || emailLower,
          relevados: stats[emailLower] || 0,
          faltantes: 0 // Si no tiene IDs asignados, no hay faltantes
        })
      }
    })

    // Ordenar por email alfabéticamente
    statsArray.sort((a, b) => a.email.localeCompare(b.email))

    console.log('Estadísticas calculadas:', statsArray)
    console.log('Total de registros:', data.length)
    
    setAdminStats(statsArray)
  }

  // Función para abrir el panel admin
  const handleAdminClick = async () => {
    calculateAdminStats()
    await loadPermissions() // Cargar permisos cuando se abre el panel
    setShowAdminPanel(true)
  }

  // Función para abrir modal de asignación de permisos
  const handleAssignPermissionsClick = (email: string) => {
    setSelectedUserForPermission(email)
    // Buscar permisos existentes del usuario
    const existingPermissions = allPermissions.find(p => p.email.toLowerCase() === email.toLowerCase())
    setPermissionIdsInput(existingPermissions?.allowedIds.join(', ') || '')
    setShowPermissionsModal(true)
  }

  // Función para guardar permisos
  const handleSavePermissions = async () => {
    if (!selectedUserForPermission) {
      alert('Error: No hay usuario seleccionado')
      return
    }

    const token = getAccessToken()
    if (!token) {
      alert('Error: No estás autenticado')
      return
    }

    // Parsear IDs separados por comas
    const ids = permissionIdsInput
      .split(',')
      .map(id => id.trim())
      .filter(id => id.length > 0)

    try {
      const response = await fetch('/api/permissions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`,
        },
        body: JSON.stringify({
          email: selectedUserForPermission,
          allowedIds: ids,
        }),
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error al guardar permisos')
      }

      // Recargar permisos
      await loadPermissions()
      
      setShowPermissionsModal(false)
      setSelectedUserForPermission('')
      setPermissionIdsInput('')
      
      alert(`✅ Permisos asignados correctamente a ${selectedUserForPermission}\n\nIDs asignados: ${ids.length > 0 ? ids.join(', ') : 'ninguno'}\n\nEl usuario solo podrá ver las filas con estos IDs en la primera columna.`)
    } catch (err: any) {
      alert(`Error al guardar permisos: ${err.message}`)
      console.error('Error guardando permisos:', err)
    }
  }

  // Funciones para edición - CORRECCIÓN: Usar ID real en lugar de índice
  const handleEditClick = (rowIndex: number) => {
    console.log('handleEditClick called with rowIndex:', rowIndex)
    console.log('filteredData length:', filteredData.length)
    console.log('filteredData[rowIndex]:', filteredData[rowIndex])

    if (!filteredData[rowIndex]) {
      console.error('No se encontró la fila en filteredData para el índice:', rowIndex)
      return
    }

    const rowData = [...filteredData[rowIndex]]
    // CORRECCIÓN: Obtener el ID real de la primera columna
    const realId = String(filteredData[rowIndex][0]).trim()
    console.log('ID real extraído:', realId)

    // Si existe la columna "fecha relevada", actualizar con la fecha actual
    if (fechaRelevadaIndex !== -1) {
      rowData[fechaRelevadaIndex] = getCurrentDate()
    }
    // Si existe la columna "relevado por", actualizar con el email del usuario
    if (relevadoPorIndex !== -1) {
      if (userEmail) {
        rowData[relevadoPorIndex] = userEmail
        console.log('Email asignado al abrir modal:', userEmail)
      } else {
        console.warn('No hay email disponible para asignar')
      }
    }
    setEditFormData(rowData)
    setEditingRowId(realId) // CORRECCIÓN: Guardar el ID real
    console.log('Modal abierto con editingRowId:', realId)
  }

  const handleEditChange = (colIndex: number, value: string) => {
    const newData = [...editFormData]
    newData[colIndex] = value
    setEditFormData(newData)
  }

  const handleCloseModal = () => {
    console.log('Cerrando modal')
    setEditingRowId(null)
    setEditFormData([])
  }

  const handleSaveEdit = async () => {
    if (!editingRowId) return

    setSaving(true)
    setError(null)

    try {
      // Encontrar la fila en los datos locales usando el ID real
      const originalRowIndex = data.findIndex((row) => {
        // Comparar por ID real (primera columna)
        return row[0] !== null && row[0] !== undefined &&
               String(row[0]).trim().toLowerCase() === editingRowId.toLowerCase()
      })

      if (originalRowIndex === -1) {
        throw new Error(`No se pudo encontrar la fila con ID ${editingRowId} en los datos locales`)
      }

      console.log('Fila encontrada en datos locales, índice:', originalRowIndex)

      // Actualizar la fecha relevada con la fecha actual antes de guardar
      const dataToSave = [...editFormData]
      
      // Asegurar que el array tenga el tamaño correcto
      while (dataToSave.length < headers.length) {
        dataToSave.push('')
      }
      
      if (fechaRelevadaIndex !== -1) {
        dataToSave[fechaRelevadaIndex] = getCurrentDate()
        console.log('Fecha actualizada:', getCurrentDate(), 'en índice:', fechaRelevadaIndex)
      }
      
      // Actualizar "relevado por" con el email del usuario antes de guardar
      if (relevadoPorIndex !== -1) {
        if (userEmail) {
          dataToSave[relevadoPorIndex] = userEmail
          console.log('Email guardado en dataToSave:', userEmail, 'en índice:', relevadoPorIndex)
          console.log('Valor antes de guardar en índice', relevadoPorIndex, ':', editFormData[relevadoPorIndex])
          console.log('Valor después de actualizar en índice', relevadoPorIndex, ':', dataToSave[relevadoPorIndex])
        } else {
          console.warn('No hay email disponible para guardar. userEmail:', userEmail)
        }
      } else {
        console.warn('Índice de "Relevado por" no encontrado. relevadoPorIndex:', relevadoPorIndex)
      }
      
      console.log('Datos completos a guardar:', dataToSave)
      console.log('Email del usuario actual:', userEmail)
      
      // Actualizar el sheet usando la API
      const token = getAccessToken()
      if (!token) {
        throw new Error('No estás autenticado')
      }

      const response = await fetch('/api/update', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`,
        },
        body: JSON.stringify({
          rowId: editingRowId, // El backend buscará la fila por este ID
          values: dataToSave,
        }),
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error al guardar')
      }

      // Actualizar los datos locales
      const newData = [...data]
      newData[originalRowIndex] = [...dataToSave]
      setData(newData)

      // Cerrar el modal
      handleCloseModal()

      // Recargar los datos para asegurar sincronización
      await loadData()
      
      // Actualizar estadísticas de admin si el panel está abierto
      if (showAdminPanel) {
        calculateAdminStats()
      }
    } catch (err: any) {
      const errorMessage = err.message || err.result?.error?.message || 'Error al guardar los cambios'
      setError(`Error al guardar: ${errorMessage}`)
    } finally {
      setSaving(false)
    }
  }

  console.log('Render component - editingRowId:', editingRowId, 'data length:', data.length, 'filteredData length:', filteredData.length)

  return (
    <div className="min-h-screen bg-gray-50 py-4 sm:py-8 px-2 sm:px-4">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="bg-white rounded-lg shadow-md p-4 sm:p-6 mb-4 sm:mb-6">
          <h1 className="text-xl sm:text-2xl md:text-3xl font-bold text-gray-800 mb-3 sm:mb-4">Clarin Censo 2026</h1>
          
          {/* Botones de autenticación */}
          <div className="flex flex-wrap gap-2 sm:gap-3 mb-4">
            {!isAuthenticated && (
      <button
        id="authorize_button"
        ref={authorizeButtonRef}
        onClick={handleAuthClick}
                className="bg-blue-600 hover:bg-blue-700 text-white px-4 sm:px-6 py-2 rounded-lg text-sm sm:text-base font-semibold transition-colors duration-200 w-full sm:w-auto"
      >
                Iniciar Sesión
      </button>
            )}
      <button
        id="signout_button"
        ref={signoutButtonRef}
        onClick={handleSignoutClick}
              className="bg-red-600 hover:bg-red-700 text-white px-4 sm:px-6 py-2 rounded-lg text-sm sm:text-base font-semibold transition-colors duration-200 w-full sm:w-auto"
              style={{ visibility: isAuthenticated ? 'visible' : 'hidden' }}
            >
              Cerrar Sesión
            </button>
            {isAuthenticated && (
              <>
                <button
                  onClick={loadData}
                  disabled={loading}
                  className="bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white px-4 sm:px-6 py-2 rounded-lg text-sm sm:text-base font-semibold transition-colors duration-200 w-full sm:w-auto"
                >
                  {loading ? 'Cargando...' : 'Actualizar'}
                </button>
                {isAdmin && (
                  <button
                    onClick={handleAdminClick}
                    className="bg-purple-600 hover:bg-purple-700 text-white px-4 sm:px-6 py-2 rounded-lg text-sm sm:text-base font-semibold transition-colors duration-200 w-full sm:w-auto"
                  >
                    ADMIN
                  </button>
                )}
              </>
            )}
          </div>

          {/* Buscador */}
          {isAuthenticated && data.length > 0 && (
            <div className="mt-4">
              <div className="flex flex-col sm:flex-row gap-3 mb-3">
                <div className="flex-1">
                  <label htmlFor="search-type" className="block text-xs sm:text-sm font-medium text-gray-700 mb-2">
                    Buscar por:
                  </label>
                  <select
                    id="search-type"
                    value={searchType}
                    onChange={(e) => setSearchType(e.target.value as 'id' | 'paquete')}
                    className="w-full px-3 sm:px-4 py-2 text-sm sm:text-base border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white"
                  >
                    <option value="id">ID</option>
                    <option value="paquete">N° Paquete</option>
                  </select>
                </div>
                <div className="flex-1">
                  <label htmlFor="search" className="block text-xs sm:text-sm font-medium text-gray-700 mb-2">
                    Término de búsqueda:
                  </label>
                  <input
                    id="search"
                    type="text"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    placeholder={
                      searchType === 'id'
                        ? "Ingresa el ID exacto (ej: 123)..."
                        : "Ingresa el paquete exacto (ej: subte)..."
                    }
                    className="w-full px-3 sm:px-4 py-2 text-sm sm:text-base border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
              </div>
              <p className="text-xs sm:text-sm text-gray-500 mt-2">
                Mostrando {filteredData.length} de {data.length} registros
                {searchTerm && (
                  <span className="block sm:inline">
                    {` (filtrados por ${searchType === 'id' ? 'ID exacto' : 'Paquete exacto'}: "${searchTerm}")`}
                  </span>
                )}
                {!isAdmin && userPermissions.allowedIds.length > 0 && (
                  <span className="block text-xs text-gray-400 mt-1">
                    Tienes acceso a {userPermissions.allowedIds.length} ID(s) permitido(s)
                  </span>
                )}
              </p>
            </div>
          )}
        </div>

        {/* Error Message */}
        {error && (
          <div className="bg-red-50 border-l-4 border-red-500 p-3 sm:p-4 mb-4 sm:mb-6 rounded">
            <div className="flex">
              <div className="ml-3">
                <p className="text-xs sm:text-sm text-red-700 break-words">{error}</p>
              </div>
            </div>
          </div>
        )}

        {/* Loading State */}
        {loading && (
          <div className="bg-white rounded-lg shadow-md p-6 sm:p-8 text-center">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
            <p className="mt-4 text-sm sm:text-base text-gray-600">Cargando datos...</p>
          </div>
        )}

        {/* Data Table */}
        {!loading && isAuthenticated && filteredData.length > 0 && (
          <>
            <div className="bg-white rounded-lg shadow-md overflow-hidden mb-4">
              <div className="overflow-x-auto -mx-2 sm:mx-0">
                <div className="inline-block min-w-full align-middle">
                  <table className="min-w-full divide-y divide-gray-200">
                    <thead className="bg-gray-50">
                      <tr>
                        <th className="px-2 sm:px-4 md:px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider sticky left-0 bg-gray-50 z-10">
                          Acciones
                        </th>
                        {headers.map((header, index) => (
                          <th
                            key={index}
                            className="px-2 sm:px-4 md:px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider"
                          >
                            <span className="hidden sm:inline">{header || `Columna ${index + 1}`}</span>
                            <span className="sm:hidden">{index + 1}</span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-gray-200">
                      {paginatedData.map((row, rowIndex) => {
                        const actualRowIndex = startIndex + rowIndex
                        // CORRECCIÓN: Buscar por ID en lugar de comparación de objetos
                        const originalRowIndex = filteredData.findIndex(r =>
                          r && row && r[0] === row[0] && String(r[0]).trim() === String(row[0]).trim()
                        )
                        console.log('Row', rowIndex, 'originalRowIndex:', originalRowIndex, 'ID:', row[0])
                        return (
                          <tr key={actualRowIndex} className="hover:bg-gray-50 transition-colors">
                            <td className="px-2 sm:px-4 md:px-6 py-3 sm:py-4 whitespace-nowrap text-xs sm:text-sm sticky left-0 bg-white z-10">
                              <button
                                onClick={() => {
                                  console.log('Botón editar clickeado, llamando handleEditClick con:', originalRowIndex)
                                  handleEditClick(originalRowIndex)
                                }}
                                className="bg-blue-600 hover:bg-blue-700 text-white px-2 sm:px-4 py-1.5 sm:py-2 rounded-lg text-xs sm:text-sm font-medium transition-colors whitespace-nowrap"
                              >
                                Editar
                              </button>
                            </td>
                            {headers.map((_, colIndex) => (
                              <td
                                key={colIndex}
                                className="px-2 sm:px-4 md:px-6 py-3 sm:py-4 text-xs sm:text-sm text-gray-900 max-w-xs sm:max-w-none"
                              >
                                <div className="truncate sm:whitespace-normal">
                                  {row[colIndex] !== null && row[colIndex] !== undefined
                                    ? String(row[colIndex])
                                    : '-'}
                                </div>
                              </td>
                            ))}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Paginación */}
            {totalPages > 1 && (
              <div className="bg-white rounded-lg shadow-md p-3 sm:p-4">
                <div className="flex flex-col sm:flex-row items-center justify-between gap-3">
                  <div className="text-xs sm:text-sm text-gray-700 text-center sm:text-left">
                    Mostrando {startIndex + 1} a {Math.min(endIndex, filteredData.length)} de {filteredData.length} registros
                  </div>
                  <div className="flex flex-wrap gap-1 sm:gap-2 justify-center">
                    <button
                      onClick={() => setCurrentPage(1)}
                      disabled={currentPage === 1}
                      className="px-2 sm:px-3 py-1.5 sm:py-2 border border-gray-300 rounded-lg text-xs sm:text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      <span className="hidden sm:inline">Primera</span>
                      <span className="sm:hidden">«</span>
                    </button>
                    <button
                      onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                      disabled={currentPage === 1}
                      className="px-2 sm:px-3 py-1.5 sm:py-2 border border-gray-300 rounded-lg text-xs sm:text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      <span className="hidden sm:inline">Anterior</span>
                      <span className="sm:hidden">‹</span>
                    </button>
                    
                    {getPageNumbers().map((page, index) => (
                      page === '...' ? (
                        <span key={`ellipsis-${index}`} className="px-2 sm:px-3 py-1.5 sm:py-2 text-gray-500 text-xs sm:text-sm">
                          ...
                        </span>
                      ) : (
                        <button
                          key={page}
                          onClick={() => setCurrentPage(page as number)}
                          className={`px-2 sm:px-3 py-1.5 sm:py-2 border rounded-lg text-xs sm:text-sm font-medium ${
                            currentPage === page
                              ? 'bg-blue-600 text-white border-blue-600'
                              : 'border-gray-300 text-gray-700 bg-white hover:bg-gray-50'
                          }`}
                        >
                          {page}
                        </button>
                      )
                    ))}
                    
                    <button
                      onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                      disabled={currentPage === totalPages}
                      className="px-2 sm:px-3 py-1.5 sm:py-2 border border-gray-300 rounded-lg text-xs sm:text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      <span className="hidden sm:inline">Siguiente</span>
                      <span className="sm:hidden">›</span>
                    </button>
                    <button
                      onClick={() => setCurrentPage(totalPages)}
                      disabled={currentPage === totalPages}
                      className="px-2 sm:px-3 py-1.5 sm:py-2 border border-gray-300 rounded-lg text-xs sm:text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      <span className="hidden sm:inline">Última</span>
                      <span className="sm:hidden">»</span>
                    </button>
                  </div>
                </div>
              </div>
            )}
          </>
        )}

        {/* Empty State */}
        {!loading && isAuthenticated && filteredData.length === 0 && data.length === 0 && !error && (
          <div className="bg-white rounded-lg shadow-md p-6 sm:p-8 text-center">
            <p className="text-sm sm:text-base text-gray-600">No hay datos para mostrar. Haz clic en "Actualizar" para cargar los datos.</p>
          </div>
        )}


        {/* No Results from Search */}
        {!loading && isAuthenticated && data.length > 0 && filteredData.length === 0 && (
          <div className="bg-white rounded-lg shadow-md p-6 sm:p-8 text-center">
            <p className="text-sm sm:text-base text-gray-600">No se encontraron resultados para "{searchTerm}" (búsqueda exacta por {searchType === 'id' ? 'ID' : 'Paquete'})</p>
            <button
              onClick={() => {
                setSearchTerm('')
                setSearchType('id')
              }}
              className="mt-4 text-sm sm:text-base text-blue-600 hover:text-blue-800 underline"
            >
              Limpiar búsqueda
            </button>
          </div>
        )}

        {/* Modal de Edición */}
        {(() => {
          console.log('Render modal, editingRowId:', editingRowId)
          return editingRowId !== null
        })() && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-2 sm:p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[95vh] sm:max-h-[90vh] overflow-y-auto m-2">
              <div className="sticky top-0 bg-white border-b border-gray-200 px-4 sm:px-6 py-3 sm:py-4 flex justify-between items-center z-10">
                <h2 className="text-lg sm:text-xl md:text-2xl font-bold text-gray-800">Editar Registro</h2>
                <button
                  onClick={handleCloseModal}
                  className="text-gray-500 hover:text-gray-700 text-2xl sm:text-3xl font-bold"
                  disabled={saving}
                >
                  ×
                </button>
              </div>
              
              <div className="p-4 sm:p-6">
                <div className="space-y-4">
                  {headers.map((header, colIndex) => {
                    const isFechaRelevada = colIndex === fechaRelevadaIndex
                    const isRelevadoPor = colIndex === relevadoPorIndex
                    const isId = colIndex === 0
                    const isDisabled = saving || isId || isFechaRelevada || isRelevadoPor
                    
                    return (
                      <div key={colIndex}>
                        <label className="block text-xs sm:text-sm font-medium text-gray-700 mb-2">
                          {header || `Columna ${colIndex + 1}`}
                        </label>
                        <input
                          type="text"
                          value={editFormData[colIndex] !== null && editFormData[colIndex] !== undefined 
                            ? String(editFormData[colIndex]) 
                            : ''}
                          onChange={(e) => handleEditChange(colIndex, e.target.value)}
                          className="w-full px-3 sm:px-4 py-2 text-sm sm:text-base border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent disabled:bg-gray-100 disabled:cursor-not-allowed"
                          disabled={isDisabled}
                        />
                        {isId && (
                          <p className="text-xs text-gray-500 mt-1">El ID no se puede modificar</p>
                        )}
                        {isFechaRelevada && (
                          <p className="text-xs text-gray-500 mt-1">La fecha se actualiza automáticamente al guardar</p>
                        )}
                        {isRelevadoPor && (
                          <p className="text-xs text-gray-500 mt-1">Se actualiza automáticamente con tu email al guardar</p>
                        )}
                      </div>
                    )
                  })}
                </div>
              </div>

              <div className="sticky bottom-0 bg-gray-50 border-t border-gray-200 px-4 sm:px-6 py-3 sm:py-4 flex flex-col sm:flex-row justify-end gap-2 sm:gap-3">
                <button
                  onClick={handleCloseModal}
                  disabled={saving}
                  className="px-4 sm:px-6 py-2 border border-gray-300 rounded-lg text-xs sm:text-sm text-gray-700 bg-white hover:bg-gray-50 font-medium disabled:opacity-50 disabled:cursor-not-allowed transition-colors w-full sm:w-auto"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleSaveEdit}
                  disabled={saving}
                  className="px-4 sm:px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs sm:text-sm font-medium disabled:opacity-50 disabled:cursor-not-allowed transition-colors w-full sm:w-auto"
                >
                  {saving ? 'Guardando...' : 'Guardar Cambios'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Panel Admin */}
        {showAdminPanel && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-2 sm:p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[95vh] sm:max-h-[90vh] overflow-y-auto m-2">
              <div className="sticky top-0 bg-white border-b border-gray-200 px-4 sm:px-6 py-3 sm:py-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 z-10">
                <h2 className="text-lg sm:text-xl md:text-2xl font-bold text-gray-800">Panel de Administración</h2>
                <div className="flex gap-2 sm:gap-3 w-full sm:w-auto">
                  <button
                    onClick={() => {
                      calculateAdminStats()
                    }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-medium transition-colors flex-1 sm:flex-none"
                  >
                    Actualizar
                  </button>
                  <button
                    onClick={() => setShowAdminPanel(false)}
                    className="text-gray-500 hover:text-gray-700 text-2xl sm:text-3xl font-bold"
                  >
                    ×
                  </button>
                </div>
              </div>
              
              <div className="p-4 sm:p-6">
                <div className="mb-4">
                  <p className="text-sm sm:text-base text-gray-600">
                    Total de registros: <span className="font-bold text-gray-800">{data.length}</span>
                  </p>
                </div>
                
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-3 sm:gap-4">
                  {adminStats.map((stat, index) => (
                    <div
                      key={index}
                      className="bg-gray-50 border-2 border-gray-200 rounded-lg p-6 hover:shadow-md transition-shadow"
                    >
                      <div className="mb-3 sm:mb-4">
                        <h3 className="text-base sm:text-lg font-semibold text-gray-700 mb-1 sm:mb-2">Usuario:</h3>
                        <p className="text-gray-900 font-mono text-xs sm:text-sm break-all">{stat.email}</p>
                      </div>
                      
                      <div className="grid grid-cols-2 gap-2 sm:gap-4">
                        <div className="bg-green-100 border border-green-300 rounded-lg p-4">
                          <p className="text-sm text-green-700 font-medium mb-1">Relevados</p>
                          <p className="text-3xl font-bold text-green-800">{stat.relevados}</p>
                        </div>
                        
                        <div className="bg-red-100 border border-red-300 rounded-lg p-4">
                          <p className="text-sm text-red-700 font-medium mb-1">Faltantes</p>
                          <p className="text-3xl font-bold text-red-800">{stat.faltantes}</p>
                        </div>
                      </div>
                      
                      {(() => {
                        const totalAsignado = stat.relevados + stat.faltantes
                        const porcentaje = totalAsignado > 0 ? (stat.relevados / totalAsignado) * 100 : 0

                        return totalAsignado > 0 && (
                          <div className="mt-4">
                            <div className="w-full bg-gray-200 rounded-full h-2">
                              <div
                                className="bg-green-600 h-2 rounded-full transition-all duration-300"
                                style={{ width: `${porcentaje}%` }}
                              ></div>
                            </div>
                            <p className="text-xs text-gray-500 mt-1 text-center">
                              {porcentaje.toFixed(1)}% completado ({stat.relevados} de {totalAsignado} asignados)
                            </p>
                          </div>
                        )
                      })()}
                      
                      <div className="mt-4 pt-4 border-t border-gray-300">
                        <button
                          onClick={() => handleAssignPermissionsClick(stat.email)}
                          className="w-full bg-purple-600 hover:bg-purple-700 text-white px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-medium transition-colors mb-3"
                        >
                          Asignar IDs Permitidos
                        </button>
                        {(() => {
                          const userPerms = allPermissions.find(p => p.email.toLowerCase() === stat.email.toLowerCase())
                          if (userPerms && userPerms.allowedIds.length > 0) {
                            return (
                              <div className="mt-3">
                                <p className="text-xs font-medium text-gray-700 mb-2">IDs asignados:</p>
                                <div className="flex flex-wrap gap-1">
                                  {userPerms.allowedIds.map((id, idx) => (
                                    <span
                                      key={idx}
                                      className="bg-purple-100 text-purple-800 px-2 py-1 rounded text-xs font-mono font-semibold"
                                    >
                                      {id}
                                    </span>
                                  ))}
                                </div>
                                <p className="text-xs text-gray-500 mt-2">
                                  Este usuario solo verá las filas con estos IDs en la primera columna
                                </p>
                              </div>
                            )
                          } else {
                            return (
                              <p className="text-xs text-orange-600 text-center mt-2">
                                ⚠️ No tiene IDs asignados (no verá ningún registro)
                              </p>
                            )
                          }
                        })()}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Modal de Asignación de Permisos */}
        {showPermissionsModal && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[60] p-2 sm:p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[95vh] sm:max-h-[90vh] overflow-y-auto m-2">
              <div className="border-b border-gray-200 px-4 sm:px-6 py-3 sm:py-4 flex justify-between items-center sticky top-0 bg-white z-10">
                <h2 className="text-lg sm:text-xl md:text-2xl font-bold text-gray-800">Asignar IDs Permitidos</h2>
                <button
                  onClick={() => {
                    setShowPermissionsModal(false)
                    setSelectedUserForPermission('')
                    setPermissionIdsInput('')
                  }}
                  className="text-gray-500 hover:text-gray-700 text-2xl sm:text-3xl font-bold"
                >
                  ×
                </button>
              </div>
              
              <div className="p-4 sm:p-6">
                <div className="mb-4">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Usuario:
                  </label>
                  <p className="text-gray-900 font-mono text-sm bg-gray-50 p-2 rounded border">
                    {selectedUserForPermission}
                  </p>
                </div>
                
                <div className="mb-4">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    IDs permitidos (separados por comas):
                  </label>
                  <textarea
                    value={permissionIdsInput}
                    onChange={(e) => setPermissionIdsInput(e.target.value)}
                    placeholder="Ejemplo: 1, 2, 3, 4, 5"
                    className="w-full px-3 sm:px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent h-32 font-mono text-xs sm:text-sm"
                  />
                  <p className="text-xs text-gray-600 mt-2 font-medium">
                    📋 Instrucciones:
                  </p>
                  <ul className="text-xs text-gray-600 mt-1 list-disc list-inside space-y-1">
                    <li>Ingresa los IDs separados por comas (ejemplo: 1, 2, 3, 4, 5)</li>
                    <li>Solo las filas con estos IDs en la <strong>primera columna</strong> serán visibles para este usuario</li>
                    <li>Si dejas vacío, el usuario no verá ningún registro</li>
                    <li>Los IDs se comparan de forma exacta (sin importar mayúsculas/minúsculas)</li>
                    <li>Los permisos se guardan en Google Sheets</li>
                  </ul>
                  {data.length > 0 && (
                    <div className="mt-3 p-3 bg-blue-50 border border-blue-200 rounded-lg">
                      <p className="text-xs font-medium text-blue-800 mb-2">
                        💡 IDs disponibles en el sheet (primeros 20):
                      </p>
                      <div className="flex flex-wrap gap-1 max-h-24 overflow-y-auto">
                        {Array.from(new Set(data.slice(0, 100).map(row => row[0]).filter(id => id !== null && id !== undefined && String(id).trim() !== ''))).slice(0, 20).map((id, idx) => (
                          <span
                            key={idx}
                            className="bg-blue-100 text-blue-800 px-2 py-1 rounded text-xs font-mono cursor-pointer hover:bg-blue-200"
                            onClick={() => {
                              const currentIds = permissionIdsInput.split(',').map(i => i.trim()).filter(i => i)
                              if (!currentIds.includes(String(id).trim())) {
                                setPermissionIdsInput([...currentIds, String(id).trim()].join(', '))
                              }
                            }}
                            title="Click para agregar"
                          >
                            {String(id).trim()}
                          </span>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {(() => {
                  const userPerms = allPermissions.find(p => p.email.toLowerCase() === selectedUserForPermission.toLowerCase())
                  if (userPerms && userPerms.allowedIds.length > 0) {
                    return (
                      <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
                        <p className="text-sm font-medium text-blue-800 mb-2">IDs actualmente asignados:</p>
                        <div className="flex flex-wrap gap-2">
                          {userPerms.allowedIds.map((id, idx) => (
                            <span
                              key={idx}
                              className="bg-blue-100 text-blue-800 px-2 py-1 rounded text-xs font-mono"
                            >
                              {id}
                            </span>
                          ))}
                        </div>
                      </div>
                    )
                  }
                  return null
                })()}
              </div>

              <div className="border-t border-gray-200 px-4 sm:px-6 py-3 sm:py-4 flex flex-col sm:flex-row justify-end gap-2 sm:gap-3 sticky bottom-0 bg-white">
                <button
                  onClick={() => {
                    setShowPermissionsModal(false)
                    setSelectedUserForPermission('')
                    setPermissionIdsInput('')
                  }}
                  className="px-4 sm:px-6 py-2 border border-gray-300 rounded-lg text-xs sm:text-sm text-gray-700 bg-white hover:bg-gray-50 font-medium transition-colors w-full sm:w-auto"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleSavePermissions}
                  className="px-4 sm:px-6 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs sm:text-sm font-medium transition-colors w-full sm:w-auto"
                >
                  Guardar Permisos
      </button>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  )
}
