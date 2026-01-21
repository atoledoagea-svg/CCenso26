import { google } from 'googleapis'

const ADMIN_EMAILS = [
  'atoledo.agea@gmail.com',
  'lmiranda.maslogistica@gmail.com',
  'emiliano.lopez.mkt@gmail.com'
]

/**
 * Valida el token de Google OAuth y obtiene el email del usuario
 */
export async function validateGoogleToken(accessToken: string): Promise<{ email: string; isAdmin: boolean } | null> {
  try {
    const oauth2Client = new google.auth.OAuth2()
    oauth2Client.setCredentials({ access_token: accessToken })

    const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client })
    const userInfo = await oauth2.userinfo.get()

    if (!userInfo.data.email) {
      return null
    }

    const email = userInfo.data.email.toLowerCase()
    const isAdmin = ADMIN_EMAILS.includes(email)

    return { email, isAdmin }
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

