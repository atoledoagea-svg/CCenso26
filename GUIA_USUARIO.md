# 📋 Guía de Usuario - Sistema de Relevamiento PDV

## 📱 Versión para Usuarios de Campo

Esta guía está diseñada para los usuarios que realizan el relevamiento de Puntos de Venta (PDV) en el campo.

---

## 📑 Índice

1. [Acceso al Sistema](#1-acceso-al-sistema)
2. [Pantalla Principal](#2-pantalla-principal)
3. [Buscar un PDV](#3-buscar-un-pdv)
4. [Editar un PDV](#4-editar-un-pdv)
5. [Estados del Puesto](#5-estados-del-puesto)
6. [Subir Fotos](#6-subir-fotos)
7. [Capturar Ubicación GPS](#7-capturar-ubicación-gps)
8. [Guardar Cambios](#8-guardar-cambios)
9. [Descargar Cuestionario PDF](#9-descargar-cuestionario-pdf)
10. [Preguntas Frecuentes](#10-preguntas-frecuentes)

---

## 1. Acceso al Sistema

### 1.1 Ingresar al Sistema

1. Abre tu navegador (Chrome, Safari, Firefox)
2. Ingresa a la dirección web proporcionada por tu supervisor
3. Verás la pantalla de bienvenida con el logo de Clarín

### 1.2 Iniciar Sesión con Google

1. Presiona el botón **"Continuar con Google"**
2. Selecciona tu cuenta de correo electrónico autorizada
3. Acepta los permisos solicitados (acceso a hojas de cálculo)

> ⚠️ **Importante:** Solo podrás acceder si tu correo fue autorizado por un administrador. Si no puedes entrar, contacta a tu supervisor.

### 1.3 Sesión Exitosa

Una vez dentro verás:
- Tu correo electrónico en la esquina superior derecha
- El contador de PDV relevados (ejemplo: "15/50 PDV")
- La tabla con los PDV asignados a tu usuario

---

## 2. Pantalla Principal

### 2.1 Elementos de la Pantalla

| Elemento | Descripción |
|----------|-------------|
| **Logo Clarín** | Identificación del sistema |
| **Tu correo** | Usuario conectado |
| **Contador PDV** | Muestra cuántos PDV has relevado del total asignado |
| **Cerrar Sesión** | Botón para salir del sistema |
| **Barra de búsqueda** | Para buscar PDV por ID o Paquete |
| **Filtro de estado** | Filtrar por relevados/sin relevar |
| **Tabla de datos** | Lista de PDV asignados |

### 2.2 La Tabla de PDV

La tabla muestra todos los PDV que tienes asignados. Cada fila tiene:

| Columna | Descripción |
|---------|-------------|
| **Acciones** | Botones ✎ (editar) y 📍 (ubicación) |
| **ID** | Número único del PDV |
| **Paquete** | Zona o grupo al que pertenece |
| **Dirección** | Ubicación del PDV |
| **Otros campos** | Información adicional del PDV |

### 2.3 Colores de las Filas

- **Blanco**: PDV pendiente de relevar
- **Verde claro**: PDV ya relevado ✓

---

## 3. Buscar un PDV

### 3.1 Buscar por ID

1. En el menú desplegable, selecciona **"ID"**
2. Escribe el número de ID exacto en el campo de búsqueda
3. La tabla mostrará solo ese PDV

### 3.2 Buscar por Paquete

1. En el menú desplegable, selecciona **"Paquete"**
2. Escribe parte del nombre del paquete
3. La tabla mostrará todos los PDV que contengan ese texto

### 3.3 Limpiar Búsqueda

- Presiona la **X** dentro del campo de búsqueda
- O borra el texto manualmente

### 3.4 Filtrar por Estado

Usa el filtro desplegable para ver:
- **📋 Todos los PDV**: Muestra todos
- **✅ Solo relevados**: Solo los que ya completaste
- **⏳ Sin relevar**: Solo los pendientes

---

## 4. Editar un PDV

### 4.1 Abrir el Formulario de Edición

1. Ubica el PDV en la tabla
2. Presiona el botón **✎** (lápiz) en la columna de acciones
3. Se abrirá el formulario de edición

> 💡 Si el PDV ya fue relevado (fila verde), el sistema te preguntará si estás seguro de querer editarlo.

### 4.2 Campos del Formulario

El formulario muestra todos los campos del PDV que puedes completar:

#### Campos Automáticos (no editables)
- **Fecha de relevamiento**: Se completa automáticamente al guardar
- **Relevado por**: Se completa con tu correo electrónico

#### Campos de Solo Lectura
- **ID**: No se puede modificar
- **Provincia**: No se puede modificar

#### Campos Obligatorios ⚠️
Están marcados con **"* Obligatorio"** y fondo amarillo:
- **Paquete**: Siempre obligatorio
- **Venta productos no editoriales**: Obligatorio si el puesto está activo
- **Teléfono**: Obligatorio si el puesto está activo (poner 0 si no se obtiene)

### 4.3 Tipos de Campos

| Tipo | Descripción |
|------|-------------|
| **Texto** | Escribe directamente |
| **Desplegable** | Selecciona una opción de la lista |
| **Desplegable con búsqueda** | Escribe para filtrar opciones (ej: Localidad) |
| **Área de texto** | Para comentarios largos |

### 4.4 Cancelar Edición

- Presiona **"Cancelar"** o la **X** en la esquina
- El sistema te pedirá confirmación antes de cerrar
- Los cambios no guardados se perderán

---

## 5. Estados del Puesto

Al abrir el formulario, lo primero que debes indicar es el **estado del puesto**:

### 5.1 ✅ Puesto Activo (Abierto)
- El puesto está funcionando normalmente
- **Debes completar todos los campos obligatorios**
- Completa la mayor cantidad de información posible

### 5.2 ❌ Cerrado Definitivamente
- El puesto cerró permanentemente
- Al seleccionar esta opción, el campo "Estado Kiosco" se marca como "Cerrado definitivamente"
- Los demás campos mantienen sus valores originales
- Solo necesitas guardar

### 5.3 ❓ No se Encontró el Puesto
- El puesto no existe en la dirección indicada
- Al seleccionar esta opción, el campo "Estado Kiosco" se marca como "No se encuentra el puesto"
- Solo necesitas guardar

### 5.4 ⚠️ Zona Peligrosa
- No es seguro acceder a la zona del puesto
- Al seleccionar esta opción, el campo "Estado Kiosco" se marca como "Zona Peligrosa"
- Solo necesitas guardar

> 💡 **¿Te equivocaste?** Si seleccionaste por error una opción de cierre, simplemente vuelve a seleccionar **"Puesto Activo"** para restaurar los campos.

---

## 6. Subir Fotos

### 6.1 Sección de Foto del PDV

Al final del formulario encontrarás la sección **"📷 Foto del PDV"**.

### 6.2 Opciones de Subida

#### En Computadora (Desktop):
- **🖼️ Galería**: Selecciona una imagen de tu computadora

#### En Celular (Móvil):
- **🖼️ Galería**: Selecciona una foto del carrete/galería
- **📷 Cámara** (con 📍 GPS): Abre la cámara para tomar una foto nueva

### 6.3 Tomar una Foto con la Cámara (Móvil)

1. Presiona el botón **"📷 Cámara"**
2. Se abrirá la cámara de tu celular
3. Toma la foto del PDV
4. Confirma la foto
5. La imagen se subirá automáticamente
6. **Se capturará tu ubicación GPS automáticamente**

### 6.4 Si Ya Tienes Coordenadas Guardadas

Cuando tomas una nueva foto con cámara y el PDV ya tiene coordenadas, el sistema te preguntará:

```
📍 Este registro ya tiene coordenadas:
Latitud: -34.584369
Longitud: -58.405580

¿Deseas actualizar las coordenadas con tu ubicación actual?
```

- **Aceptar**: Actualiza las coordenadas con tu ubicación actual
- **Cancelar**: Sube la foto pero conserva las coordenadas anteriores

### 6.5 Quitar una Foto

Si ya subiste una foto y quieres cambiarla:
1. Presiona el botón **"✕ Quitar imagen"**
2. Luego sube una nueva foto

### 6.6 Formatos Permitidos

- **Formatos**: JPG, PNG, GIF
- **Tamaño máximo**: 32MB

---

## 7. Capturar Ubicación GPS

### 7.1 Métodos para Capturar Ubicación

Hay **dos formas** de guardar la ubicación GPS de un PDV:

#### Método 1: Con la Foto (Cámara)
- Al tomar una foto con el botón **"📷 Cámara"**, la ubicación se captura automáticamente
- Las coordenadas se guardan cuando presionas "Guardar Cambios"

#### Método 2: Botón de Ubicación Rápida 📍
- En la tabla principal, cada fila tiene un botón **📍**
- Al presionarlo, captura tu ubicación actual y la guarda directamente
- No necesitas abrir el formulario de edición

### 7.2 Usar el Botón de Ubicación Rápida 📍

1. Ubica el PDV en la tabla
2. Presiona el botón **📍** (al lado del botón de editar ✎)
3. El sistema obtendrá tu ubicación GPS
4. Se guardará automáticamente en el Excel

> ⚠️ Si el PDV ya tiene coordenadas, el sistema te preguntará si quieres sobrescribirlas.

### 7.3 Permisos de Ubicación

Para que funcione la captura GPS:
1. Tu navegador te pedirá permiso para acceder a la ubicación
2. Debes presionar **"Permitir"**
3. Asegúrate de tener el GPS activado en tu celular

### 7.4 Si No Funciona la Ubicación

Si no se puede obtener la ubicación:
- Verifica que el GPS esté activado
- Verifica que hayas dado permiso al navegador
- Intenta en un lugar con mejor señal
- Reinicia el navegador si es necesario

---

## 8. Guardar Cambios

### 8.1 Antes de Guardar

Verifica que:
1. El estado del puesto sea correcto
2. Los campos obligatorios estén completos (marcados en amarillo)
3. La información sea correcta

### 8.2 Proceso de Guardado

1. Presiona el botón **"Guardar Cambios"**
2. El sistema te pedirá confirmación
3. Presiona **"Aceptar"**
4. Espera a que aparezca el mensaje de éxito

### 8.3 Errores al Guardar

Si falta algún campo obligatorio, verás un mensaje como:

```
⚠️ Por favor complete los siguientes campos obligatorios:

- Paquete
- Venta productos no editoriales
- Teléfono (poner 0 si no se obtiene)
```

Completa los campos indicados y vuelve a guardar.

### 8.4 Después de Guardar

- El formulario se cerrará automáticamente
- La fila del PDV se pondrá de color **verde** ✓
- El contador de PDV relevados se actualizará
- Los datos se guardan directamente en la hoja de cálculo

---

## 9. Descargar Cuestionario PDF

Si necesitas llenar un cuestionario a mano (para PDV nuevos o sin conexión):

### 9.1 Cómo Descargar

1. Presiona el botón **"➕ Nuevo PDV"** en la barra superior
2. Selecciona **"📋 CUESTIONARIO PDF"**
3. Se abrirá una ventana con el formulario para imprimir
4. Imprime o guarda como PDF

### 9.2 Contenido del Cuestionario

El cuestionario incluye todos los campos necesarios:
- Datos de ubicación
- Estado del puesto
- Características del puesto
- Datos de contacto
- Espacio para observaciones

### 9.3 Uso del Cuestionario

1. Imprime el cuestionario antes de salir al campo
2. Llénalo a mano durante la visita
3. Luego ingresa los datos al sistema
4. O entrega el cuestionario a tu supervisor

---

## 10. Preguntas Frecuentes

### ❓ ¿Por qué no puedo ver ningún PDV?

Tu cuenta necesita tener IDs asignados. Contacta a tu administrador para que te asigne PDV.

### ❓ ¿Puedo editar un PDV que ya relevé?

Sí, pero el sistema te pedirá confirmación ya que los datos serán sobrescritos.

### ❓ ¿Qué pasa si cierro el navegador sin guardar?

Los cambios se perderán. **Siempre guarda antes de cerrar.**

### ❓ ¿Puedo ver los PDV de otros usuarios?

No, solo puedes ver y editar los PDV asignados a tu cuenta.

### ❓ ¿Qué hago si el teléfono del PDV no está disponible?

Escribe **0** en el campo de teléfono.

### ❓ ¿Puedo usar el sistema en mi celular?

Sí, el sistema está optimizado para dispositivos móviles. Puedes:
- Ver y editar PDV
- Tomar fotos con la cámara
- Capturar ubicación GPS

### ❓ ¿Cómo sé cuántos PDV me faltan?

Mira el contador en la parte superior: **"X/X PDV relevados"**
- Primer número: Completados
- Segundo número: Total asignado

### ❓ ¿Por qué algunos campos están deshabilitados?

- Si seleccionaste "Cerrado", "No encontrado" o "Zona Peligrosa", algunos campos se deshabilitan
- Los campos de fecha y relevador siempre se llenan automáticamente
- El campo Provincia y el ID no son editables

### ❓ ¿Puedo trabajar sin conexión a internet?

No, el sistema requiere conexión a internet para funcionar. Puedes usar el cuestionario PDF para llenar a mano y luego cargar los datos cuando tengas conexión.

### ❓ ¿Las fotos se guardan en mi celular?

No, las fotos se suben directamente al servidor y se guarda solo el enlace en el Excel.

### ❓ ¿Por qué no aparece el botón de Cámara en mi computadora?

El botón de Cámara solo aparece en dispositivos móviles. En computadora solo puedes adjuntar imágenes desde la galería.

### ❓ ¿Las coordenadas GPS se guardan automáticamente?

- **Con la cámara**: Sí, se capturan automáticamente al tomar la foto
- **Con galería**: No, las fotos de galería no capturan ubicación
- **Botón 📍**: Puedes guardar ubicación sin tomar foto

### ❓ ¿Qué hago si me equivoqué en un dato?

Simplemente vuelve a editar el PDV, corrige el dato y guarda los cambios.

### ❓ ¿Cómo cierro sesión?

Presiona el botón **"Cerrar Sesión"** en la esquina superior derecha.

---

## 📞 Soporte

Si tienes problemas técnicos o dudas que no están en esta guía, contacta a tu supervisor o al equipo de soporte.

---

*Última actualización: Febrero 2026*



