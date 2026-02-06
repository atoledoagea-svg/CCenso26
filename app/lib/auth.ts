import { google } from 'googleapis'

// Tipos de roles de usuario
export type UserRole = 'user' | 'supervisor' | 'admin'

// Tipo de nivel de usuario (desde la hoja de permisos)
export type UserLevel = 1 | 2 | 3

// Super administradores (siempre tienen nivel 3, aunque no estén en la hoja de permisos)
const SUPER_ADMIN_EMAILS = [
  'atoledo.agea@gmail.com',
  'lmiranda.maslogistica@gmail.com',
  'emiliano.lopez.mkt@gmail.com'
]

/**
 * Verifica si un email es super admin (siempre tiene acceso completo)
 */
export function isSuperAdmin(email: string): boolean {
  return SUPER_ADMIN_EMAILS.includes(email.toLowerCase())
}

/**
 * Convierte el nivel numérico (1, 2, 3) a rol de usuario
 * Nivel 1 = user (Usuario normal)
 * Nivel 2 = supervisor (Acceso intermedio, sin GPS de usuarios)
 * Nivel 3 = admin (Acceso completo)
 */
export function levelToRole(level: UserLevel): UserRole {
  switch (level) {
    case 3: return 'admin'
    case 2: return 'supervisor'
    default: return 'user'
  }
}

/**
 * Obtiene el rol del usuario según su email y nivel de permisos
 * Los super admins siempre tienen rol 'admin' independientemente del nivel en la hoja
 */
export function getUserRole(email: string, level: UserLevel = 1): UserRole {
  const emailLower = email.toLowerCase()
  if (SUPER_ADMIN_EMAILS.includes(emailLower)) {
    return 'admin'
  }
  return levelToRole(level)
}

/**
 * Valida el token de Google OAuth y obtiene el email del usuario
 * El rol se determina después, basado en el nivel de la hoja de permisos
 */
export async function validateGoogleToken(accessToken: string): Promise<{ email: string; isSuperAdmin: boolean } | null> {
  try {
    const oauth2Client = new google.auth.OAuth2()
    oauth2Client.setCredentials({ access_token: accessToken })

    const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client })
    const userInfo = await oauth2.userinfo.get()

    if (!userInfo.data.email) {
      return null
    }

    const email = userInfo.data.email.toLowerCase()
    
    return { 
      email, 
      isSuperAdmin: isSuperAdmin(email)
    }
  } catch (error) {
    console.error('Error validando token de Google:', error)
    return null
  }
}

/**
 * Obtiene el token de acceso del header Authorization
 */
export function getAccessTokenFromRequest(request: Request): string | null {
  const authHeader = request.headers.get('authorization')
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return null
  }
  return authHeader.substring(7)
}

