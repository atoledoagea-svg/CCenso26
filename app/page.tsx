'use client'

import { useEffect, useState, useCallback } from 'react'
import Image from 'next/image'

declare global {
  interface Window {
    gapi: any
    google: any
  }
}

interface SheetData {
  headers: string[]
  data: any[][]
  permissions: {
    allowedIds: string[]
    isAdmin: boolean
  }
}

interface Permission {
  email: string
  allowedIds: string[]
}

export default function Home() {
  const [tokenClient, setTokenClient] = useState<any>(null)
  const [gapiInited, setGapiInited] = useState(false)
  const [gisInited, setGisInited] = useState(false)
  
  const [isAuthorized, setIsAuthorized] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const [userEmail, setUserEmail] = useState<string | null>(null)
  const [accessToken, setAccessToken] = useState<string | null>(null)
  
  // Dashboard states
  const [sheetData, setSheetData] = useState<SheetData | null>(null)
  const [loadingData, setLoadingData] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [editingRow, setEditingRow] = useState<number | null>(null)
  const [editedValues, setEditedValues] = useState<any[]>([])
  const [originalValues, setOriginalValues] = useState<any[]>([]) // Valores originales del Excel
  const [saving, setSaving] = useState(false)
  const [searchTerm, setSearchTerm] = useState('')
  const [searchType, setSearchType] = useState<'id' | 'paquete'>('id')
  const [filterRelevado, setFilterRelevado] = useState<'todos' | 'relevados' | 'no_relevados'>('todos')
  const [showStats, setShowStats] = useState(false)
  
  // Admin panel states
  const [showPermissionsPanel, setShowPermissionsPanel] = useState(false)
  const [allPermissions, setAllPermissions] = useState<Permission[]>([])
  const [editingPermission, setEditingPermission] = useState<Permission | null>(null)
  const [newPermIds, setNewPermIds] = useState('')
  const [savingPermissions, setSavingPermissions] = useState(false)
  const [newUserEmail, setNewUserEmail] = useState('')
  
  // Pagination states
  const [currentPage, setCurrentPage] = useState(1)
  const itemsPerPage = 50
  
  // Puesto Activo/Cerrado state
  const [puestoStatus, setPuestoStatus] = useState<'abierto' | 'cerrado' | 'no_encontrado' | 'zona_peligrosa' | ''>('')

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

  const loadSheetData = useCallback(async (token: string) => {
    setLoadingData(true)
    setError(null)
    
    try {
      const response = await fetch('/api/data', {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      })
      
      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error cargando datos')
      }
      
      const data = await response.json()
      setSheetData(data)
      
      if (data.permissions?.isAdmin) {
        loadPermissions(token)
      }
    } catch (err: any) {
      setError(err.message)
    } finally {
      setLoadingData(false)
    }
  }, [])

  const loadPermissions = async (token: string) => {
    try {
      const response = await fetch('/api/permissions', {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      })
      
      if (response.ok) {
        const data = await response.json()
        if (data.permissions) {
          setAllPermissions(data.permissions)
        }
      }
    } catch (err) {
      console.error('Error cargando permisos:', err)
    }
  }

  function handleAuthClick() {
    if (!tokenClient) return

    setIsLoading(true)

    tokenClient.callback = async (resp: any) => {
      if (resp.error !== undefined) {
        setIsLoading(false)
        console.error('Auth error:', resp)
        return
      }
      
      const token = window.gapi.client.getToken().access_token
      setAccessToken(token)
      setIsAuthorized(true)
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
      
      await loadSheetData(token)
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
      setIsAuthorized(false)
      setUserEmail(null)
      setAccessToken(null)
      setSheetData(null)
      setAllPermissions([])
    }
  }

  const handleEditRow = (rowIndex: number) => {
    if (sheetData) {
      const rowData = [...sheetData.data[rowIndex]]
      
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
      setPuestoStatus('abierto')
    }
  }

  const handleCancelEdit = () => {
    const confirmCancel = window.confirm('¿Estás seguro de que deseas cancelar? Los cambios no guardados se perderán.')
    if (confirmCancel) {
      setEditingRow(null)
      setEditedValues([])
      setOriginalValues([])
      setPuestoStatus('')
    }
  }

  // Find column indexes for auto-fill fields
  const getAutoFillIndexes = () => {
    if (!sheetData) return { fechaIndex: -1, relevadorIndex: -1 }
    
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    
    // Find "Fecha" column (could be "fecha", "fecha de relevo", etc.)
    const fechaIndex = headers.findIndex(h => 
      h.includes('fecha') || h === 'date' || h.includes('fecha de relevo')
    )
    
    // Find "Relevador por:" column
    const relevadorIndex = headers.findIndex(h => 
      h.includes('relevador') || h.includes('relevado por') || h.includes('censado por')
    )
    
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
  const handlePuestoStatusChange = (status: 'abierto' | 'cerrado' | 'no_encontrado' | 'zona_peligrosa') => {
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
      const camposCerradoIndexes = getCamposCerradoIndexes()
      const newValues = [...editedValues]
      
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = 'Puesto Cerrado DEFINITIVAMENTE'
      })
      
      setEditedValues(newValues)
    } else if (status === 'no_encontrado') {
      const camposCerradoIndexes = getCamposCerradoIndexes()
      const newValues = [...editedValues]
      
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = 'NO SE ENCONTRO PUESTO'
      })
      
      setEditedValues(newValues)
    } else if (status === 'zona_peligrosa') {
      const camposCerradoIndexes = getCamposCerradoIndexes()
      const newValues = [...editedValues]
      
      camposCerradoIndexes.forEach(idx => {
        newValues[idx] = 'ZONA PELIGROSA'
      })
      
      setEditedValues(newValues)
    }
  }

  const handleSaveRow = async () => {
    if (!accessToken || editingRow === null || !sheetData || !userEmail) return
    
    // Validar campos obligatorios solo si el puesto está activo
    if (puestoStatus === 'abierto') {
      const headers = sheetData.headers.map(h => h.toLowerCase().trim())
      
      // Buscar índice de Venta productos no editoriales
      const ventaNoEditorialIndex = headers.findIndex(h => 
        h.includes('venta') && h.includes('no editorial')
      )
      
      // Buscar índice de Teléfono
      const telefonoIndex = headers.findIndex(h => 
        h.includes('telefono') || h.includes('teléfono')
      )
      
      const errores: string[] = []
      
      if (ventaNoEditorialIndex !== -1 && !String(editedValues[ventaNoEditorialIndex] || '').trim()) {
        errores.push('- Venta productos no editoriales')
      }
      
      if (telefonoIndex !== -1 && !String(editedValues[telefonoIndex] || '').trim()) {
        errores.push('- Teléfono (poner 0 si no se obtiene)')
      }
      
      if (errores.length > 0) {
        alert(`⚠️ Por favor complete los siguientes campos obligatorios:\n\n${errores.join('\n')}`)
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
      const valuesToSave = [...editedValues]
      
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
      
      // Set user email in relevador field
      if (relevadorIndex !== -1) {
        valuesToSave[relevadorIndex] = userEmail
      }
      
      const response = await fetch('/api/update', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          rowId: rowId,
          values: valuesToSave
        })
      })
      
      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error guardando datos')
      }
      
      const newData = [...sheetData.data]
      newData[editingRow] = valuesToSave
      setSheetData({ ...sheetData, data: newData })
      setEditingRow(null)
      setEditedValues([])
    } catch (err: any) {
      alert('Error: ' + err.message)
    } finally {
      setSaving(false)
    }
  }

  const handleSavePermissions = async (email: string, ids: string[]) => {
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
          allowedIds: ids
        })
      })
      
      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error guardando permisos')
      }
      
      await loadPermissions(accessToken)
      setEditingPermission(null)
      setNewPermIds('')
      setNewUserEmail('')
    } catch (err: any) {
      alert('Error: ' + err.message)
    } finally {
      setSavingPermissions(false)
    }
  }

  // Calculate stats for a user based on "Relevador por:" field
  const getUserStats = (userEmailPerm: string, allowedIds: string[]) => {
    if (!sheetData) return { relevados: 0, faltantes: 0, total: 0 }
    
    const totalAssigned = allowedIds.length
    
    // Find the "Relevador por:" column index
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    const relevadorIndex = headers.findIndex(h => 
      h.includes('relevador') || h.includes('relevado por') || h.includes('censado por')
    )
    
    if (relevadorIndex === -1) {
      // If column not found, fallback to checking if ID exists
      const existingIds = sheetData.data.map(row => String(row[0] || '').toLowerCase())
      const relevados = allowedIds.filter(id => existingIds.includes(id.toLowerCase())).length
      return { relevados, faltantes: totalAssigned - relevados, total: totalAssigned }
    }
    
    // Count IDs where "Relevador por:" matches the user's email
    const userEmailLower = userEmailPerm.toLowerCase()
    const allowedIdsLower = allowedIds.map(id => id.toLowerCase())
    
    let relevados = 0
    for (const row of sheetData.data) {
      const rowId = String(row[0] || '').toLowerCase()
      const relevadorEmail = String(row[relevadorIndex] || '').toLowerCase().trim()
      
      // Check if this row's ID is in the user's allowed IDs
      // AND if the relevador field contains the user's email
      if (allowedIdsLower.includes(rowId) && relevadorEmail === userEmailLower) {
        relevados++
      }
    }
    
    const faltantes = totalAssigned - relevados
    
    return { relevados, faltantes, total: totalAssigned }
  }

  // Get Paquete column index
  const getPaqueteIndex = () => {
    if (!sheetData) return -1
    const headers = sheetData.headers.map(h => h.toLowerCase().trim())
    return headers.findIndex(h => h.includes('paquete'))
  }

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

  // Descargar estadísticas como CSV/Excel
  const downloadStatsAsExcel = () => {
    const stats = getFieldStats()
    if (stats.length === 0) return
    
    // Obtener conteo de relevados
    const { relevadorIndex } = getAutoFillIndexes()
    const totalRelevados = sheetData?.data.filter(row => {
      if (relevadorIndex === -1) return false
      return String(row[relevadorIndex] || '').trim() !== ''
    }).length || 0
    
    let csvContent = '\uFEFF' // BOM para UTF-8 en Excel
    
    // Título y resumen
    csvContent += 'ESTADÍSTICAS DE PDV RELEVADOS\n'
    csvContent += `Fecha de generación:,${new Date().toLocaleDateString('es-AR')}\n`
    csvContent += `Total PDV:,${sheetData?.data.length || 0}\n`
    csvContent += `PDV Relevados:,${totalRelevados}\n`
    csvContent += `PDV Pendientes:,${(sheetData?.data.length || 0) - totalRelevados}\n`
    csvContent += '\n'
    
    // Datos de cada campo
    stats.forEach(fieldStat => {
      const total = fieldStat.data.reduce((sum, d) => sum + d.count, 0)
      
      csvContent += `${fieldStat.fieldName}\n`
      csvContent += 'Opción,Cantidad,Porcentaje\n'
      
      fieldStat.data.forEach(d => {
        const percent = Math.round((d.count / total) * 100)
        // Escapar comas en las etiquetas
        const label = d.label.includes(',') ? `"${d.label}"` : d.label
        csvContent += `${label},${d.count},${percent}%\n`
      })
      
      csvContent += `Total,${total},100%\n`
      csvContent += '\n'
    })
    
    // Crear y descargar el archivo
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

  const filteredData = sheetData?.data.filter(row => {
    // Filtro por búsqueda de texto
    if (searchTerm.trim()) {
      const searchValue = searchTerm.trim().toLowerCase()
      
      if (searchType === 'id') {
        // Exact match for ID (first column)
        const rowId = String(row[0] || '').toLowerCase().trim()
        if (rowId !== searchValue) return false
      } else {
        // Flexible match for Paquete column (contains)
        const paqueteIndex = getPaqueteIndex()
        if (paqueteIndex !== -1) {
          const paqueteValue = String(row[paqueteIndex] || '').toLowerCase().trim()
          if (!paqueteValue.includes(searchValue)) return false
        }
      }
    }
    
    // Filtro por estado de relevamiento
    if (filterRelevado !== 'todos') {
      const { relevadorIndex } = getAutoFillIndexes()
      const relevadorValue = relevadorIndex !== -1 ? String(row[relevadorIndex] || '').trim() : ''
      const isRelevado = relevadorValue !== ''
      
      if (filterRelevado === 'relevados' && !isRelevado) return false
      if (filterRelevado === 'no_relevados' && isRelevado) return false
    }
    
    return true
  }) || []

  // Pagination logic
  const totalPages = Math.ceil(filteredData.length / itemsPerPage)
  const startIndex = (currentPage - 1) * itemsPerPage
  const endIndex = startIndex + itemsPerPage
  const paginatedData = filteredData.slice(startIndex, endIndex)
  
  useEffect(() => {
    setCurrentPage(1)
  }, [searchTerm, searchType, filterRelevado])

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

  // Login Screen
  if (!isAuthorized) {
    return (
      <div className="login-container">
        <div className="login-left">
          <div className="welcome-content">
            <div className="logo-wrapper">
              <Image 
                src="/Clarín_logo.svg.png" 
                alt="Clarín Logo" 
                width={180} 
                height={50}
                priority
              />
            </div>
            <h1>Bienvenido a</h1>
            <h2>Relevamiento de PDV</h2>
            <p>
              Accede con tu cuenta de Google para consultar y gestionar los datos del relevamiento de manera segura.
            </p>
          </div>
          
          <div className="clouds-decoration">
            <div className="cloud cloud-1"></div>
            <div className="cloud cloud-2"></div>
            <div className="cloud cloud-3"></div>
          </div>
        </div>

        <div className="login-right">
          <div className="login-card">
            <h3>Iniciar Sesión</h3>
            <p>Usa tu cuenta de Google para acceder al sistema</p>
            
            <button 
              className={`google-btn ${isLoading ? 'loading' : ''}`}
              onClick={handleAuthClick}
              disabled={!isReady || isLoading}
            >
              {isLoading ? (
                <div className="spinner"></div>
              ) : (
                <svg viewBox="0 0 24 24">
                  <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
                  <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
                  <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/>
                  <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
                </svg>
              )}
              {isLoading ? 'Conectando...' : 'Continuar con Google'}
            </button>
            
            {!isReady && (
              <div style={{ textAlign: 'center', marginTop: '1rem', color: 'var(--gray-300)' }}>
                <small>Cargando...</small>
              </div>
            )}
          </div>
          
          <div className="footer-text">
            <p>© 2026 Clarín - Todos los derechos reservados</p>
          </div>
        </div>
      </div>
    )
  }

  // Dashboard Screen
  return (
    <div className="dashboard">
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
                  </div>
                  {puestoStatus === 'cerrado' && (
                    <div className="puesto-cerrado-notice">
                      <span className="notice-icon">⚠️</span>
                      <span>Los campos relevantes se han rellenado automáticamente con "Puesto Cerrado DEFINITIVAMENTE".</span>
                    </div>
                  )}
                  {puestoStatus === 'no_encontrado' && (
                    <div className="puesto-cerrado-notice puesto-no-encontrado-notice">
                      <span className="notice-icon">❓</span>
                      <span>Los campos relevantes se han rellenado automáticamente con "NO SE ENCONTRO PUESTO".</span>
                    </div>
                  )}
                  {puestoStatus === 'zona_peligrosa' && (
                    <div className="puesto-cerrado-notice puesto-peligrosa-notice">
                      <span className="notice-icon">🚨</span>
                      <span>Los campos relevantes se han rellenado automáticamente con "ZONA PELIGROSA".</span>
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
                    
                    // Campos obligatorios destacados
                    const isCampoObligatorio = isVentaNoEditorialField || isTelefonoField
                    
                    const estadoKioscoOptions = [
                      'Abierto',
                      'Cerrado ahora',
                      'Abre ocasionalmente',
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
                    const camposCerradoIndexes = getCamposCerradoIndexes()
                    const isCampoCerrado = (puestoStatus === 'cerrado' || puestoStatus === 'no_encontrado' || puestoStatus === 'zona_peligrosa') && camposCerradoIndexes.includes(idx)
                    
                    return (
                      <div key={idx} className={`edit-field ${isAutoField ? 'auto-field' : ''} ${isCampoCerrado ? 'campo-cerrado' : ''} ${isCampoObligatorio && puestoStatus === 'abierto' ? 'campo-obligatorio' : ''}`}>
                        <label>
                          {isSugerenciasField ? displayHeader : header}
                          {isAutoField && <span className="auto-badge">Auto</span>}
                          {isCampoObligatorio && puestoStatus === 'abierto' && <span className="obligatorio-badge">* Obligatorio</span>}
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
                        {isEstadoKioscoField || isDiasAtencionField || isHorarioField || isEscaparateField || isUbicacionField || isFachadaField || isVentaNoEditorialField || isRepartoField || isSuscripcionesField || isParadaOnlineField || isMayorVentaField || isDistribuidoraField ? (() => {
                          const currentOptions = isDiasAtencionField ? diasAtencionOptions : isHorarioField ? horarioOptions : isEscaparateField ? escaparateOptions : isUbicacionField ? ubicacionOptions : isFachadaField ? fachadaOptions : isVentaNoEditorialField ? ventaNoEditorialOptions : isRepartoField ? repartoOptions : isSuscripcionesField ? suscripcionesOptions : isParadaOnlineField ? paradaOnlineOptions : isMayorVentaField ? mayorVentaOptions : isDistribuidoraField ? distribuidoraOptions : estadoKioscoOptions
                          const currentValue = editedValues[idx] || ''
                          const valueExistsInOptions = currentOptions.includes(currentValue) || currentValue === ''
                          
                          return (
                            <select
                              value={currentValue}
                              onChange={(e) => {
                                if (!isCampoCerrado) {
                                  const newValues = [...editedValues]
                                  newValues[idx] = e.target.value
                                  setEditedValues(newValues)
                                }
                              }}
                              className="estado-kiosco-select"
                              disabled={isCampoCerrado}
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
                          />
                        ) : (
                          <input
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
                            className={`${isAutoField ? 'auto-input' : ''} ${isCampoObligatorio && puestoStatus === 'abierto' ? 'input-obligatorio' : ''}`}
                          />
                        )}
                      </div>
                    )
                  })}
                </div>
              </div>
              <div className="modal-footer">
                <button className="btn-secondary" onClick={handleCancelEdit} disabled={saving}>
                  Cancelar
                </button>
                <button className="btn-primary" onClick={handleSaveRow} disabled={saving}>
                  {saving ? 'Guardando...' : 'Guardar Cambios'}
                </button>
              </div>
            </div>
          </div>
        )
      })()}

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
                      await loadSheetData(accessToken)
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
                <span>Total de registros: <strong>{sheetData.data.length}</strong></span>
              </div>

              {/* Add new user form */}
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

              <div className="users-grid">
                {allPermissions.map((perm, idx) => {
                  const stats = getUserStats(perm.email, perm.allowedIds)
                  const completionPercent = stats.total > 0 
                    ? Math.round((stats.relevados / stats.total) * 100) 
                    : 0

                  return (
                    <div key={idx} className="user-card">
                      <div className="user-card-header">
                        <span className="user-label">Usuario:</span>
                        <span className="user-card-email">{perm.email}</span>
                      </div>
                      
                      <div className="stats-boxes">
                        <div className="stat-box stat-green">
                          <span className="stat-label">Relevados</span>
                          <span className="stat-number">{stats.relevados}</span>
                        </div>
                        <div className="stat-box stat-red">
                          <span className="stat-label">Faltantes</span>
                          <span className="stat-number">{stats.faltantes}</span>
                        </div>
                      </div>

                      <div className="progress-section">
                        <div className="progress-bar">
                          <div 
                            className="progress-fill" 
                            style={{ width: `${completionPercent}%` }}
                          ></div>
                        </div>
                        <span className="progress-text">
                          {completionPercent}% completado ({stats.relevados} de {stats.total} asignados)
                        </span>
                      </div>

                      <button 
                        className="btn-assign"
                        onClick={() => {
                          setEditingPermission(perm)
                          setNewPermIds(perm.allowedIds.join(', '))
                        }}
                      >
                        Asignar IDs Permitidos
                      </button>

                      <div className="ids-section">
                        <span className="ids-label">IDs asignados:</span>
                        <div className="ids-tags">
                          {perm.allowedIds.length === 0 ? (
                            <span className="no-ids">Sin IDs asignados</span>
                          ) : (
                            perm.allowedIds.map((id, i) => (
                              <span key={i} className="id-tag">{id}</span>
                            ))
                          )}
                        </div>
                        <span className="ids-hint">
                          Este usuario solo verá las filas con estos IDs en la primera columna
                        </span>
                      </div>
                    </div>
                  )
                })}

                {allPermissions.length === 0 && (
                  <div className="no-users">
                    <p>No hay usuarios con permisos configurados</p>
                    <p className="hint">Agrega un usuario usando el formulario de arriba</p>
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
          <div className="stats-panel" onClick={e => e.stopPropagation()}>
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
            <div className="stats-body">
              <div className="stats-summary">
                <div className="summary-card">
                  <span className="summary-number">{sheetData.data.length}</span>
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
                <div className="summary-card summary-orange">
                  <span className="summary-number">
                    {sheetData.data.filter(row => {
                      const { relevadorIndex } = getAutoFillIndexes()
                      return relevadorIndex === -1 || String(row[relevadorIndex] || '').trim() === ''
                    }).length}
                  </span>
                  <span className="summary-label">Pendientes</span>
                </div>
              </div>
              
              <div className="charts-grid">
                {getFieldStats().map((fieldStat, idx) => {
                  const total = fieldStat.data.reduce((sum, d) => sum + d.count, 0)
                  let cumulativePercent = 0
                  
                  // Calcular gradiente cónico para el gráfico de pastel
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
            </div>
          </div>
        </div>
      )}

      {/* Edit Permission Modal */}
      {editingPermission && (
        <div className="modal-overlay" onClick={() => setEditingPermission(null)}>
          <div className="modal-content modal-small" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Asignar IDs</h2>
              <button className="modal-close" onClick={() => setEditingPermission(null)}>×</button>
            </div>
            <div className="modal-body">
              <p className="modal-user-email">{editingPermission.email}</p>
              <div className="edit-field">
                <label>IDs permitidos (separados por coma)</label>
                <textarea
                  value={newPermIds}
                  onChange={(e) => setNewPermIds(e.target.value)}
                  placeholder="Ej: 1, 10, 5, 8, 99, 112"
                  rows={4}
                />
              </div>
            </div>
            <div className="modal-footer">
              <button 
                className="btn-secondary" 
                onClick={() => setEditingPermission(null)} 
                disabled={savingPermissions}
              >
                Cancelar
              </button>
              <button 
                className="btn-primary" 
                onClick={() => {
                  const ids = newPermIds.split(',').map(id => id.trim()).filter(id => id)
                  handleSavePermissions(editingPermission.email, ids)
                }}
                disabled={savingPermissions}
              >
                {savingPermissions ? 'Guardando...' : 'Guardar'}
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
          {sheetData?.permissions?.isAdmin && (
            <>
              <button 
                className="btn-stats"
                onClick={() => setShowStats(true)}
              >
                📊 Estadísticas
              </button>
              <button 
                className="btn-admin"
                onClick={() => setShowPermissionsPanel(true)}
              >
                Gestionar Permisos
              </button>
            </>
          )}
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
            {sheetData?.permissions?.isAdmin && (
              <span className="admin-badge">Admin</span>
            )}
          </div>
          <button className="btn-secondary" onClick={handleSignoutClick}>
            Cerrar Sesión
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="dashboard-main">
        {/* Toolbar */}
        <div className="toolbar">
          <div className="toolbar-left">
            <div className="search-group">
              <select 
                value={searchType}
                onChange={(e) => setSearchType(e.target.value as 'id' | 'paquete')}
                className="search-select"
              >
                <option value="id">ID</option>
                <option value="paquete">Paquete</option>
              </select>
              <input
                type="text"
                placeholder={searchType === 'id' ? 'Buscar por ID exacto...' : 'Buscar paquetes que contengan...'}
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
              onChange={(e) => setFilterRelevado(e.target.value as 'todos' | 'relevados' | 'no_relevados')}
              className="filter-relevado-select"
            >
              <option value="todos">📋 Todos los PDV</option>
              <option value="relevados">✅ Solo relevados</option>
              <option value="no_relevados">⏳ Sin relevar</option>
            </select>
            <button 
              className="btn-primary"
              onClick={() => accessToken && loadSheetData(accessToken)}
              disabled={loadingData}
            >
              {loadingData ? 'Cargando...' : 'Recargar Datos'}
            </button>
          </div>
          <div className="toolbar-right">
            {sheetData && filteredData.length > 0 && (
              <span className="results-count">
                Mostrando {startIndex + 1} a {Math.min(endIndex, filteredData.length)} de {filteredData.length} registros
              </span>
            )}
          </div>
        </div>

        {/* Error Message */}
        {error && (
          <div className="error-message">
            {error}
          </div>
        )}

        {/* Loading State */}
        {loadingData && (
          <div className="loading-container">
            <div className="spinner-large"></div>
            <p>Cargando datos...</p>
          </div>
        )}

        {/* Data Table */}
        {!loadingData && sheetData && (
          <div className="table-container">
            {filteredData.length === 0 ? (
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
                    {sheetData.headers.map((header, idx) => (
                      <th key={idx}>{header}</th>
                    ))}
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
                          <button 
                            className="btn-edit"
                            onClick={() => handleEditRow(originalIndex)}
                          >
                            ✎
                          </button>
                        </td>
                        {row.map((cell, cellIdx) => (
                          <td key={cellIdx}>{String(cell || '')}</td>
                        ))}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* Pagination */}
        {!loadingData && sheetData && filteredData.length > 0 && (
          <div className="pagination-container">
            <div className="pagination-info">
              Mostrando {startIndex + 1} a {Math.min(endIndex, filteredData.length)} de {filteredData.length} registros
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
    </div>
  )
}
