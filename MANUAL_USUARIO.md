# 📋 Manual de Usuario - Sistema de Relevamiento PDV

## Índice
1. [Introducción](#introducción)
2. [Acceso al Sistema](#acceso-al-sistema)
3. [Panel Principal (Dashboard)](#panel-principal-dashboard)
4. [Búsqueda y Filtros](#búsqueda-y-filtros)
5. [Edición de PDV](#edición-de-pdv)
6. [Estados del Puesto](#estados-del-puesto)
7. [Campos del Formulario](#campos-del-formulario)
8. [Guardar Cambios](#guardar-cambios)
9. [Preguntas Frecuentes](#preguntas-frecuentes)

---

## Introducción

El Sistema de Relevamiento PDV es una herramienta diseñada para registrar y actualizar información sobre los Puntos de Venta (PDV) de Clarín. Este manual te guiará paso a paso en el uso del sistema.

---

## Acceso al Sistema

### Paso 1: Ingresar al sitio
1. Abre tu navegador web (Chrome, Firefox, Edge, Safari)
2. Ingresa a la dirección del sistema proporcionada por tu supervisor

### Paso 2: Iniciar sesión con Google
1. Haz clic en el botón **"Continuar con Google"**
2. Se abrirá una ventana de Google
3. Selecciona tu cuenta de correo electrónico autorizada
4. Acepta los permisos solicitados

> ⚠️ **Importante:** Solo podrás acceder si tu correo electrónico ha sido autorizado por un administrador.

### Paso 3: Verificación exitosa
Una vez autorizado, verás:
- Tu correo electrónico en la parte superior derecha
- El contador de PDV asignados y relevados
- La tabla con los PDV que debes relevar

---

## Panel Principal (Dashboard)

### Elementos del Dashboard

| Elemento | Descripción |
|----------|-------------|
| **Logo Clarín** | Identificación del sistema |
| **Tu correo** | Muestra el usuario conectado |
| **Contador X/X PDV** | Muestra cuántos PDV has relevado del total asignado |
| **Botón Cerrar Sesión** | Para salir del sistema |
| **Tabla de datos** | Lista de PDV asignados a tu usuario |

### La Tabla de Datos

La tabla muestra las siguientes columnas:
- **Acciones**: Botón para editar cada PDV
- **ID**: Número único de identificación del puesto
- **Paquete**: Zona o grupo al que pertenece
- **Dirección**: Ubicación del PDV
- **Y más campos** según configuración

#### Filas de colores:
- **Blanco**: PDV pendiente de relevar
- **Verde claro**: PDV ya relevado ✓

---

## Búsqueda y Filtros

### Buscar PDV

1. **Por ID**: 
   - Selecciona "Buscar por ID" en el menú desplegable
   - Escribe el número de ID en el campo de búsqueda

2. **Por Paquete**:
   - Selecciona "Buscar por Paquete" en el menú desplegable
   - Escribe el nombre del paquete

3. Para limpiar la búsqueda, haz clic en la **X** dentro del campo

### Filtrar por estado de relevamiento

Usa el filtro desplegable para mostrar:
- **Todos los PDV**: Muestra todos tus PDV asignados
- **Solo relevados**: Muestra solo los PDV que ya completaste
- **Sin relevar**: Muestra solo los PDV pendientes

### Navegación por páginas

Si tienes muchos PDV, usa los botones de paginación:
- **« Primera**: Ir a la primera página
- **‹ Anterior**: Ir a la página anterior
- **Números**: Ir a una página específica
- **Siguiente ›**: Ir a la siguiente página
- **Última »**: Ir a la última página

---

## Edición de PDV

### Abrir el formulario de edición

1. Ubica el PDV que deseas editar en la tabla
2. Haz clic en el botón **"✏️ Editar"** en la columna de acciones
3. Se abrirá el formulario de edición

> 💡 **Nota:** Si el PDV ya fue relevado (fila verde), el sistema te preguntará si estás seguro de querer editarlo nuevamente.

### Cerrar sin guardar

- Haz clic en **"Cancelar"** o en la **X** de la esquina
- El sistema te pedirá confirmación antes de cerrar

---

## Estados del Puesto

Al abrir el formulario de edición, lo primero que debes indicar es el **estado del puesto**. Tienes 4 opciones:

### ✓ Puesto Activo
- El puesto está abierto y funcionando
- Debes completar **todos los campos** del formulario
- Los campos de **Venta productos no editoriales** y **Teléfono** son **obligatorios**

### ✗ Puesto Cerrado DEFINITIVAMENTE
- El puesto cerró permanentemente
- Al seleccionar esta opción, **se rellenan automáticamente** varios campos con "Puesto Cerrado DEFINITIVAMENTE"
- Solo necesitas guardar

### ? No se encontró el puesto
- El puesto no existe en la dirección indicada
- Al seleccionar esta opción, **se rellenan automáticamente** varios campos con "NO SE ENCONTRO PUESTO"
- Solo necesitas guardar

### ⚠ Zona Peligrosa
- No es seguro acceder a la zona del puesto
- Al seleccionar esta opción, **se rellenan automáticamente** varios campos con "ZONA PELIGROSA"
- Solo necesitas guardar

> 💡 **¿Te equivocaste?** Si seleccionaste por error "Cerrado", "No encontrado" o "Zona Peligrosa", simplemente vuelve a hacer clic en **"Puesto Activo"** para restaurar los datos originales del Excel.

---

## Campos del Formulario

### Campos automáticos (no editables)
Estos campos se llenan automáticamente al guardar:
- **Fecha de relevamiento**: Se pone la fecha actual
- **Relevado por**: Se pone tu correo electrónico

### Campos de solo lectura
- **ID**: No se puede modificar
- **Provincia**: No se puede modificar

### Campos obligatorios ⚠️
Estos campos están destacados en **amarillo** y son obligatorios cuando el puesto está activo:
- **Venta productos no editoriales**: Debes seleccionar Sí o No
- **Teléfono**: Debes ingresar el número. Si no lo puedes obtener, escribe **0**

### Campos con opciones desplegables

#### Estado Kiosco
| Opción |
|--------|
| Abierto |
| Cerrado ahora |
| Abre ocasionalmente |
| Cerrado definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Días de atención
| Opción |
|--------|
| Todos los dias |
| De L a V |
| Sabado y Domingo |
| 3 veces por semana |
| 4 veces por Semana |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Horario
| Opción |
|--------|
| Mañana |
| Mañana y Tarde |
| Tarde |
| Solo reparto/Susc. |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Escaparate
| Opción |
|--------|
| Chico |
| Mediano |
| Grande |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Ubicación
| Opción |
|--------|
| Avenida |
| Barrio |
| Estación Subte/Tren |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Fachada puesto
| Opción |
|--------|
| Malo |
| Regular |
| Bueno |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Venta productos no editoriales ⚠️ (Obligatorio)
| Opción |
|--------|
| Nada |
| Poco |
| Mucho |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### ¿Se reparten los materiales Clarín?
| Opción |
|--------|
| Si |
| No |
| Ocasionalmente |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Suscripciones
| Opción |
|--------|
| Si |
| No |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### ¿Utiliza Parada Online?
| Opción |
|--------|
| Si |
| No |
| No sabe |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Mayor venta
| Opción |
|--------|
| Mostrador |
| Reparto |
| Suscripciones |
| No sabe / No comparte |
| Puesto Cerrado |
| Cerrado Definitivamente |
| Zona Peligrosa |
| No se encuentra el puesto |

#### Distribuidora
| Opción |
|--------|
| Barracas |
| Belgrano |
| Barrio Norte |
| Zunni |
| Recova |
| Boulogne |
| Del Parque |
| Roca/La Boca |
| Lavalle |
| Mariano Acosta |
| Nueva Era |
| San Isidro |
| Ex Rubbo |
| Ex Lugano |
| Ex Jose C Paz |

### Campo de texto libre
- **Sugerencias del PDV**: Aquí puedes escribir comentarios, observaciones o sugerencias del punto de venta. Es opcional.

---

## Guardar Cambios

### Antes de guardar

1. Verifica que el estado del puesto sea correcto
2. Si es "Puesto Activo", asegúrate de completar:
   - **Venta productos no editoriales** (obligatorio)
   - **Teléfono** (obligatorio - poner 0 si no se obtiene)
3. Completa todos los campos que puedas

### Guardar

1. Haz clic en el botón **"Guardar Cambios"**
2. El sistema te pedirá confirmación: **"¿Estás seguro de guardar los cambios?"**
3. Haz clic en **Aceptar** para confirmar
4. Espera a que aparezca el mensaje de éxito

### Después de guardar

- La fila del PDV se pondrá de color **verde** indicando que ya fue relevado
- El contador de PDV relevados se actualizará
- Los datos se guardan directamente en la hoja de cálculo

---

## Preguntas Frecuentes

### ¿Por qué no puedo ver ningún PDV?
Tu cuenta necesita tener IDs asignados. Contacta a tu administrador.

### ¿Puedo editar un PDV que ya relevé?
Sí, pero el sistema te pedirá confirmación ya que los datos serán sobrescritos.

### ¿Qué pasa si cierro el navegador sin guardar?
Los cambios se perderán. Siempre guarda antes de cerrar.

### ¿Puedo ver los PDV de otros usuarios?
No, solo puedes ver y editar los PDV asignados a tu cuenta.

### ¿Qué hago si el teléfono del PDV no está disponible?
Escribe **0** en el campo de teléfono.

### ¿Puedo usar el sistema en mi celular?
Sí, el sistema es responsive y funciona en dispositivos móviles. Puedes deslizar horizontalmente para ver todas las columnas de la tabla.

### ¿Cómo sé cuántos PDV me faltan?
Mira el contador en la parte superior: **X/X PDV relevados**. El primer número son los completados, el segundo el total asignado.

### ¿Por qué algunos campos están deshabilitados?
- Si seleccionaste "Cerrado", "No encontrado" o "Zona Peligrosa", los campos se deshabilitan automáticamente
- Los campos de fecha y relevador siempre están deshabilitados (se llenan automáticamente)
- El campo Provincia y el ID no son editables

### ¿Puedo trabajar sin conexión a internet?
No, el sistema requiere conexión a internet para funcionar y guardar los datos.

---

## Soporte

Si tienes problemas con el sistema, contacta a tu supervisor o administrador.

---

*Manual actualizado: Enero 2026*
*Sistema de Relevamiento PDV - Clarín*

