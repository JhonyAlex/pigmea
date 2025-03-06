# pigmea
# Sistema de Registro de Producción PIGMEA

Sistema integral de control y registro de producción desarrollado en VBA para Excel, que maneja múltiples procesos de producción: Laminación, Impresión y Rebobinación.

## Descripción General

Este sistema permite gestionar y registrar la producción diaria de diferentes procesos de manufactura, incluyendo:
- Laminación de materiales (Bicapa, No tejido, Tricapa, Antivaho)
- Procesos de Impresión 
- Procesos de Rebobinación

## Módulos del Sistema

### 1. Registro de Laminación
- Control de pedidos y metros producidos
- Gestión de diferentes tipos de laminados:
  - Bicapa
  - No tejido
  - Tricapa
  - Antivaho
- Seguimiento por turno y operario
- Control automático de semanas

### 2. Registro de Impresión
- Control de cambios de tinta
- Registro de camisas
- Control de transparencias
- Seguimiento de metros producidos
- Control de pedidos diarios

### 3. Registro de Rebobinación
- Control de pedidos y bandas
- Gestión de refilado
- Control de tipos:
  - Monolámina
  - Bicapa
  - Tricapa
- Registro de barras y micro

## Características Comunes

### Gestión de Datos
- Base de datos centralizada para cada proceso
- Seguimiento semanal automático
- Control de duplicados con opciones de:
  - Editar registro existente
  - Duplicar con nuevos valores
  - Cancelar operación

### Validaciones
- Verificación de campos obligatorios
- Validación de fechas
- Control de duplicados
- Validación de días de la semana

### Funcionalidades de Usuario
- Limpieza automática de formularios
- Sistema de respaldo y recuperación
- Manejo de errores
- Interfaz intuitiva con botones de control

## Estructura de Base de Datos

### Laminación
| Columna | Descripción |
|---------|-------------|
| A | Fecha |
| B | Semana del Año |
| C | Día |
| D | Turno |
| E | Operario |
| ... | ... |

### Impresión
| Columna | Descripción |
|---------|-------------|
| A | Fecha |
| B | Semana |
| C | Día |
| ... | ... |

### Rebobinación
| Columna | Descripción |
|---------|-------------|
| A | Fecha |
| B | Número de Pedido |
| C | Día |
| ... | ... |

## Uso del Sistema

### Registro de Datos
1. Seleccionar el módulo correspondiente (Laminación/Impresión/Rebobinación)
2. Ingresar fecha, turno y operario
3. Llenar los datos de producción específicos
4. Guardar el registro

### Control de Duplicados
En caso de encontrar un número de pedido duplicado, el sistema ofrece tres opciones:
- Editar el registro existente
- Crear un nuevo registro duplicado
- Cancelar la operación

### Limpieza de Formularios
- Botón dedicado para limpiar formularios
- Opción de mantener datos de cabecera
- Confirmación de seguridad antes de limpiar

## Formato de Registros
```
Fecha: dd/mm/yyyy
Tiempo: HH:MM:SS
Usuario: [Nombre del Operario]
```

## Desarrollador
- **Autor**: JhonyAlex
- **Última Actualización**: 2025-03-06 16:52:00
- **Versión**: 1.0

## Requisitos Técnicos
- Microsoft Excel con soporte para VBA
- Macros habilitadas
- Permisos de escritura en el directorio de trabajo
