# Guía de Despliegue - Clarin Censo 2026

## Opción 1: Vercel (Recomendado para Next.js)

### Paso 1: Preparar el proyecto
1. Asegúrate de que tu código esté en GitHub, GitLab o Bitbucket
2. Verifica que tengas un archivo `.gitignore` con:
   ```
   node_modules/
   .next/
   .env.local
   .env*.local
   ```

### Paso 2: Variables de Entorno
1. Ve a [vercel.com](https://vercel.com) y crea una cuenta
2. En el dashboard, crea un nuevo proyecto
3. Conecta tu repositorio de Git
4. En "Environment Variables", agrega:
   - No necesitas agregar nada manualmente si usas OAuth del cliente
   - Las credenciales están en el código del cliente (lo cual es normal para OAuth público)

### Paso 3: Configuración
1. Framework Preset: Next.js (debería detectarlo automáticamente)
2. Build Command: `npm run build` (automático)
3. Output Directory: `.next` (automático)
4. Install Command: `npm install` (automático)

### Paso 4: Desplegar
1. Haz clic en "Deploy"
2. Vercel construirá y desplegará automáticamente
3. Obtendrás una URL como: `tu-proyecto.vercel.app`

### Paso 5: Dominio Personalizado (Opcional)
1. Ve a Settings > Domains
2. Agrega tu dominio personalizado
3. Configura los DNS según las instrucciones

---

## Opción 2: Netlify

### Paso 1: Preparar
1. Sube tu código a Git (GitHub, GitLab, Bitbucket)

### Paso 2: Desplegar en Netlify
1. Ve a [netlify.com](https://netlify.com) y crea cuenta
2. Click en "Add new site" > "Import an existing project"
3. Conecta tu repositorio
4. Configura:
   - Build command: `npm run build`
   - Publish directory: `.next`
   - Base directory: (dejar vacío)

### Paso 3: Variables de Entorno
1. Site settings > Build & deploy > Environment
2. Agrega variables si las necesitas

---

## Opción 3: Railway

### Paso 1: Instalar Railway CLI (opcional) o usar Web UI
1. Ve a [railway.app](https://railway.app)
2. Crea una cuenta
3. Nuevo proyecto > Deploy from GitHub repo

### Paso 2: Configuración
1. Selecciona tu repositorio
2. Railway detectará Next.js automáticamente
3. Configura las variables de entorno si es necesario

---

## Opción 4: Render

### Paso 1: Crear cuenta
1. Ve a [render.com](https://render.com)
2. Crea una cuenta

### Paso 2: Nuevo Web Service
1. Conecta tu repositorio
2. Configura:
   - Build Command: `npm install && npm run build`
   - Start Command: `npm start`
   - Environment: Node

---

## ⚠️ IMPORTANTE: Configurar Google OAuth para Producción

### 1. Google Cloud Console
1. Ve a [Google Cloud Console](https://console.cloud.google.com)
2. Selecciona tu proyecto
3. APIs & Services > Credentials
4. Edita tu OAuth 2.0 Client ID

### 2. Authorized JavaScript Origins
Agrega:
- `https://tu-dominio.vercel.app` (o el dominio que uses)
- `https://tu-dominio.com` (si usas dominio personalizado)

### 3. Authorized Redirect URIs
Agrega:
- `https://tu-dominio.vercel.app` (o el dominio que uses)
- `https://tu-dominio.com` (si usas dominio personalizado)

### 4. Actualizar en el código (si es necesario)
Las credenciales ya están en el código. Si necesitas cambiarlas:
- `app/page.tsx` - líneas con `CLIENT_ID` y `API_KEY`

---

## Verificación Post-Despliegue

1. ✅ Verifica que el sitio carga correctamente
2. ✅ Prueba el botón "Iniciar Sesión"
3. ✅ Verifica que puedes autenticarte con Google
4. ✅ Prueba cargar datos
5. ✅ Verifica que los permisos funcionan

---

## Solución de Problemas Comunes

### Error: redirect_uri_mismatch
- Verifica que agregaste la URL correcta en Google Cloud Console
- Asegúrate de que la URL no tenga trailing slash (/)

### Error: CORS
- Vercel/Netlify maneja CORS automáticamente
- Si usas otro hosting, verifica la configuración

### Error: API routes no funcionan
- Verifica que estás usando Next.js 14+ (App Router)
- Asegúrate de que las rutas están en `app/api/`

---

## Recomendación Final

**Vercel es la mejor opción** porque:
- ✅ Propiedad de los creadores de Next.js
- ✅ Configuración automática
- ✅ SSL gratuito
- ✅ CDN global
- ✅ Despliegues instantáneos desde Git
- ✅ Preview deployments para cada PR

