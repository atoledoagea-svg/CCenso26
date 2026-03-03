'use client'

import { useState, useCallback } from 'react'

export type UserLevel = 1 | 2 | 3

export interface Permission {
  email: string
  allowedIds: string[]
  assignedSheet?: string
  level?: UserLevel
}

export function usePermissions(
  accessToken: string | null,
  onSessionExpired: () => void
) {
  const [permissions, setPermissions] = useState<Permission[]>([])

  const loadPermissions = useCallback(async (token: string) => {
    try {
      const response = await fetch('/api/permissions', {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      })
      if (response.status === 401) {
        onSessionExpired()
        return
      }
      if (response.ok) {
        const data = await response.json()
        if (data.permissions) {
          setPermissions(data.permissions)
        }
      }
    } catch (err) {
      console.error('Error cargando permisos:', err)
    }
  }, [onSessionExpired])

  return { permissions, loadPermissions, setPermissions }
}
