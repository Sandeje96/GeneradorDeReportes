# CLAUDE.md — Generador de Reportes PDF de Ventas (Grupo Petri)

## Descripción del Proyecto

App web con **Streamlit** que permite a un cliente subir un PDF de reporte "ABC de Productos" (generado por el sistema de gestión de Grupo Petri) y obtener automáticamente un reporte visual interactivo + exportación a Excel profesional.

## Estructura del PDF de Entrada

El PDF tiene esta estructura fija (49 páginas, ~2300 productos para la sucursal Central):

### Encabezado (se repite en cada página, IGNORAR al extraer datos)
```
Abc de Productos                           25/03/2026
                                           17:48
Fecha desde:    Fecha hasta:    Sucursal:
1/1/2026        31/1/2026       Central
Rubro y SubRubro:  Familia:    Vendedor:
Todos              Todas       Todos
...
```

### Tabla de datos (esto es lo que hay que extraer)
Columnas fijas en cada página:
| Columna | Tipo | Descripción |
|---------|------|-------------|
| Código | int/str | Código interno del producto |
| Descripción | str | Nombre del producto |
| Unidades | float | Cantidad vendida (puede tener decimales, ej: 16,026.71) |
| Costo | float | Costo total en pesos argentinos |
| Precio | float | Precio de venta total en pesos argentinos |
| Rentab. | float | Rentabilidad en pesos (Precio - Costo) |
| Marg. | str/float | Margen porcentual (ej: "81.00%") — puede ser negativo |
| Part. | str/float | Participación porcentual sobre el total de ventas |

### Pie de página (IGNORAR)
```
Productor: dariol    Hoja X de 49    ID: 0702
```

### Última página — Resumen final
Al final del PDF hay un bloque de totales:
```
Costos:         $ 109,835,007.16
Precios:        $ 296,918,053.10
Rentabilidad:   $ 187,083,045.94
Margenes:       170.33 %
```
Estos valores sirven para validar que la extracción fue correcta.

### Datos importantes del encabezado a extraer
- **Fecha desde / Fecha hasta**: período del reporte
- **Sucursal**: nombre de la sucursal
- Estos datos se usan para titular el reporte generado

## Stack Tecnológico

- **Python 3.10+**
- **pdfplumber**: extracción de tablas del PDF (NO usar tabula, NO usar camelot)
- **pandas**: procesamiento y análisis de datos
- **Streamlit**: interfaz web
- **openpyxl**: generación de Excel con formato profesional
- **plotly** o **matplotlib**: gráficos (plotly preferido para Streamlit por su interactividad)

## Estructura de Archivos

```
proyecto/
├── app.py                 # App principal Streamlit
├── extractor.py           # Módulo de extracción PDF → DataFrame
├── analyzer.py            # Módulo de análisis y cálculos
├── excel_report.py        # Módulo de generación Excel
├── requirements.txt       # Dependencias
├── README.md              # Instrucciones de uso
└── sample/                # Carpeta para PDFs de prueba
    └── central.pdf
```

## Instrucciones de Extracción del PDF

### Lógica de extracción (extractor.py)

1. Abrir el PDF con `pdfplumber`
2. Iterar página por página
3. En cada página, usar `page.extract_table()` o `page.extract_tables()` para obtener las filas
4. **Filtrar el encabezado repetido**: las primeras filas de cada página son el header del reporte (Fecha desde, Sucursal, etc.). La fila de encabezado de la tabla empieza con: `Código | Descripción | Unidades | Costo | Precio | Rentab. | Marg. | Part.`
5. **Filtrar el pie de página**: la última fila de cada página contiene "Productor:" y "Hoja X de Y"
6. **Manejar descripciones multilínea**: algunos productos tienen nombres largos que se cortan en 2 líneas en el PDF. Ejemplo:
   ```
   11885  PAPEL HIGIENICO FOFINHO HOJA SIMPLE 30M 4    90.00    52.67%   0.04%
          ROLLOS
          73,080.00   111,570.25   38,490.25
   ```
   Hay que detectar estas filas partidas y unirlas con la fila anterior.
7. **Extraer metadatos del encabezado** (solo de la primera página): Fecha desde, Fecha hasta, Sucursal
8. **Extraer totales de la última página**: Costos, Precios, Rentabilidad, Margenes

### Limpieza de datos

- Convertir strings numéricos a float: quitar comas de miles, manejar paréntesis como negativos `(11,191.11)` → `-11191.11`
- Convertir porcentajes string a float: "81.00%" → 81.00
- Manejar valores con costo 0 (ej: PIZZA LIBRE POR PERSONA tiene costo 0.00)
- Eliminar filas con unidades = 0 o productos sin datos
- El DataFrame final debe tener estas columnas: `codigo, descripcion, unidades, costo, precio, rentabilidad, margen, participacion`

### Validación

Después de extraer, sumar la columna `costo` y `precio` y comparar con los totales de la última página. Si la diferencia es mayor al 1%, mostrar un warning al usuario.

## Reporte en Streamlit (app.py)

### Layout de la app

1. **Sidebar**: upload del PDF + info del reporte (período, sucursal)
2. **Página principal** con secciones:

### Sección 1: Resumen Ejecutivo (tarjetas/métricas)
- Total Ventas (Precio): formateado como moneda AR
- Total Costos
- Rentabilidad Total
- Margen Global (%)
- Cantidad de productos
- Producto más vendido (por unidades)
- Producto más rentable (por rentabilidad $)

### Sección 2: Top 15 Productos Más Vendidos (por unidades)
- Tabla con los 15 productos
- **Gráfico de barras horizontal** con los 15 productos (eje Y = descripción, eje X = unidades vendidas)
- Colores profesionales, labels legibles

### Sección 3: Top 15 Productos por Rentabilidad ($)
- Tabla con los 15 productos que generaron más rentabilidad en pesos
- **Gráfico de barras horizontal** (eje Y = descripción, eje X = rentabilidad en $)

### Sección 4: Top 15 Productos por Precio de Venta Total
- Tabla + gráfico de barras con los que más facturaron

### Sección 5: Distribución de Participación (Pareto)
- Gráfico de barras + línea acumulativa mostrando que el 20% de los productos genera el 80% de las ventas
- Indicar cuántos productos representan el 80% de la facturación

### Sección 6: Productos con Margen Negativo (Alerta)
- Tabla de productos donde el margen es negativo (están vendiendo a pérdida)
- Destacar en rojo
- Si no hay productos con margen negativo, mostrar mensaje positivo

### Sección 7: Análisis por Rango de Margen
- Gráfico de torta/dona mostrando distribución de productos por rangos de margen:
  - Menos de 0% (pérdida)
  - 0% a 30%
  - 30% a 60%
  - 60% a 100%
  - Más de 100%

### Sección 8: Tabla completa de datos
- Tabla interactiva con todos los productos
- Con buscador/filtro
- Ordenable por cualquier columna

### Sección 9: Exportar a Excel
- Botón "Descargar Reporte Excel"
- El Excel debe contener múltiples hojas (ver abajo)

## Generación del Excel (excel_report.py)

Usar **openpyxl** para generar un Excel profesional con estas hojas:

### Hoja 1: "Resumen"
- Título: "Reporte de Ventas — [Sucursal] — [Fecha desde] a [Fecha hasta]"
- Métricas generales en formato de tarjetas
- Logo o encabezado estilizado no es necesario, pero sí formato profesional (fuente Arial, bordes, colores)

### Hoja 2: "Datos Completos"
- Tabla completa con todos los productos
- Headers con fondo azul oscuro (#1B3A5C) y texto blanco
- Filas alternadas con color de fondo suave para legibilidad
- Formato numérico correcto (moneda, porcentaje)
- Columnas autoajustadas al contenido

### Hoja 3: "Top Vendidos"
- Top 15 por unidades vendidas

### Hoja 4: "Top Rentabilidad"
- Top 15 por rentabilidad en $

### Hoja 5: "Margen Negativo"
- Productos con margen < 0, resaltados en rojo

### Hoja 6: "Análisis Pareto"
- Datos del análisis 80/20

### Gráficos en Excel
- Insertar gráficos de barras directamente en las hojas correspondientes usando openpyxl.chart
- Gráfico de barras en "Top Vendidos" y "Top Rentabilidad"

## Reglas de Formato y Estilo

### Colores del reporte
- Azul oscuro principal: `#1B3A5C`
- Dorado acento: `#C8A84E`
- Verde positivo: `#27AE60`
- Rojo negativo/alerta: `#E74C3C`
- Fondo gris claro alterno: `#F8F9FA`

### Formato de moneda
- Los valores están en pesos argentinos
- Usar formato: `$ #,##0.00` (con punto para decimales, coma para miles)
- En Streamlit mostrar con `f"$ {valor:,.2f}"`

### Formato de porcentajes
- Mostrar con un decimal: `f"{valor:.1f}%"`

## Manejo de Errores

- Si el PDF no tiene la estructura esperada (no es un ABC de Productos), mostrar error claro
- Si la extracción de tablas falla en alguna página, loguear y continuar con las demás
- Si hay productos duplicados (mismo código), sumarlos
- Manejar PDFs de otras sucursales (la estructura es la misma, solo cambia el nombre de la sucursal)

## Testing

- Probar con el archivo `sample/central.pdf` que tiene 49 páginas
- Verificar que los totales extraídos coincidan con:
  - Costos: $109,835,007.16
  - Precios: $296,918,053.10
  - Rentabilidad: $187,083,045.94
  - Margen: 170.33%
- Verificar que los productos con margen negativo se detecten (ej: PREPIZZA FUGAZZETA tiene margen -1.29%)
- Verificar que las descripciones multilínea se unan correctamente

## Dependencias (requirements.txt)

```
streamlit>=1.30.0
pdfplumber>=0.10.0
pandas>=2.0.0
openpyxl>=3.1.0
plotly>=5.18.0
xlsxwriter>=3.1.0
```

## Comando para ejecutar

```bash
streamlit run app.py
```

## Notas Adicionales

- El sistema del cliente puede generar estos PDFs para diferentes sucursales y períodos. La app debe ser genérica y funcionar con cualquier PDF que siga esta estructura.
- No hardcodear nombres de sucursales ni fechas.
- La extracción con pdfplumber es la parte más delicada. Si `extract_table()` no funciona bien, probar con `extract_text()` y parsear con regex línea por línea como fallback.
- Priorizar que la extracción sea robusta sobre que sea elegante.
