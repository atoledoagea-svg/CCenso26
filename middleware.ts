import { NextResponse } from 'next/server'
import type { NextRequest } from 'next/server'

const MAPA_PATH = '/mapa'
const COOKIE_NAME = 'clarin_session'

export function middleware(request: NextRequest) {
  if (request.nextUrl.pathname !== MAPA_PATH) {
    return NextResponse.next()
  }
  const hasSession = request.cookies.get(COOKIE_NAME)?.value === '1'
  if (!hasSession) {
    return NextResponse.redirect(new URL('/', request.url))
  }
  return NextResponse.next()
}

export const config = {
  matcher: ['/mapa'],
}
