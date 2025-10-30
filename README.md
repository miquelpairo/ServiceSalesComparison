# 📊 Comparador de Ventas por Periodo

Aplicación web desarrollada con Streamlit para comparar ventas de servicios entre dos periodos diferentes.

## 🎯 Propósito

Esta herramienta permite analizar y comparar datos de ventas exportados desde Power BI, identificando:

- ✅ Productos/servicios vendidos en ambos periodos
- ✅ Productos únicos en cada periodo
- ✅ Diferencias en cantidades e importes
- ✅ Análisis detallado por cliente, producto y representante de ventas

## 🚀 Instalación

### Requisitos Previos
- Python 3.7 o superior
- Pip (gestor de paquetes de Python)

### Instalación de Dependencias

```bash
pip install -r requirements.txt
```

El archivo `requirements.txt` incluye:
- `streamlit` - Framework de la aplicación web
- `pandas` - Procesamiento de datos
- `openpyxl` - Exportación a Excel

## 💻 Uso

### Ejecutar la Aplicación

```bash
streamlit run Servicecomparer.py
```

La aplicación se abrirá automáticamente en tu navegador en `http://localhost:8501`

### Despliegue en Streamlit Community Cloud

1. Sube el código a GitHub
2. Ve a [share.streamlit.io](https://share.streamlit.io/)
3. Conecta tu repositorio de GitHub
4. Selecciona `Servicecomparer.py` como archivo principal
5. ¡Despliega!

## 📋 Guía Paso a Paso

### Paso 1: Exportar Datos desde Power BI

1. Abre tu dashboard de ventas en Power BI
2. Selecciona la tabla de ventas que deseas analizar
3. Haz clic en **"Exportar datos"** → **"Datos subyacentes"**
4. Guarda el archivo como `.xlsx` o `.csv`

**Importante:** Exporta dos archivos, uno para cada periodo que deseas comparar.

### Paso 2: Preparar los Archivos

Asegúrate de que tus archivos de Power BI contengan las siguientes columnas:

| Columna Power BI | Descripción | Tipo |
|-----------------|-------------|------|
| `Date` | Fecha de la transacción | Fecha |
| `Business Partner Name` | Nombre del cliente | Texto |
| `ItemIdAndName` | Identificador y nombre del producto | Texto |
| `ProductType` | Tipo de producto (Servicio, Producto, etc.) | Texto |
| `Qty` | Cantidad vendida | Numérico |
| `EUR` | Importe total en euros | Numérico |
| `SalesRepresentative` | Representante de ventas | Texto |
| `Set` | Agrupación del producto | Texto |
| `Productline` | Línea de producto | Texto |

**Ejemplo de datos válidos:**
```csv
Date,Business Partner Name,ItemIdAndName,ProductType,Qty,EUR,SalesRepresentative,Set,Productline
2024-01-15,Cliente ABC,PROD001 - Producto A,Servicio,10,1500.00,Juan Pérez,Set A,Línea 1
2024-01-16,Cliente XYZ,PROD002 - Producto B,Producto,5,750.50,María García,Set B,Línea 2
```

### Paso 3: Cargar Archivos en la Aplicación

1. **Nombra tus periodos:**
   - Periodo 1: Por ejemplo, "Q1 2024" o "Enero 2024"
   - Periodo 2: Por ejemplo, "Q1 2023" o "Enero 2023"

2. **Sube los archivos:**
   - Haz clic en el botón de carga del Periodo 1
   - Haz clic en el botón de carga del Periodo 2

3. **Vista previa:** Revisa que los datos se hayan cargado correctamente

### Paso 4: Configurar Columnas (si es necesario)

La aplicación detecta automáticamente las columnas estándar de Power BI. Si tus columnas tienen nombres diferentes:

1. Ve a la sección **"🛠️ Asignación de columnas"**
2. Usa los selectores desplegables para mapear cada campo
3. Verifica que el mapeo sea correcto antes de continuar

### Paso 5: Aplicar Filtros

#### Filtro por Fecha 📅
- Ajusta el rango de fechas para cada periodo
- Por defecto, se selecciona todo el rango disponible
- Útil para comparar meses específicos dentro de un año

#### Filtro por Tipo de Producto 🎯
- Selecciona qué tipos de productos incluir en el análisis
- Ejemplo: Incluir solo "Servicio" y "Producto"
- Excluir tipos como "Alquiler" o "Muestra"

### Paso 6: Revisar Métricas

Antes de descargar, revisa el **resumen de la comparativa**:

- 💰 **Total Periodo 1:** Suma total de ventas del primer periodo
- 💰 **Total Periodo 2:** Suma total de ventas del segundo periodo
- 📊 **Diferencia Total:** Variación entre ambos periodos
- 🔄 **Registros comunes:** Ventas que aparecen en ambos periodos
- 📈 **Total registros:** Cantidad total de combinaciones únicas

### Paso 7: Descargar Resultado

1. Haz clic en **"📥 Descargar comparativa en Excel"**
2. El archivo se descargará con el nombre: `comparativa_[P1]_vs_[P2].xlsx`
3. Abre el archivo en Excel para análisis detallado

## 📊 Estructura del Archivo Excel

El archivo generado contiene **5 hojas** con diferentes vistas:

### Hoja 1: "Comparativa"
Tabla principal con todas las combinaciones cliente-producto:
- Cantidades e importes de ambos periodos
- Diferencias calculadas automáticamente
- Vista completa para análisis global

### Hoja 2: "Datos Originales"
Todas las transacciones originales con columna "Origen":
- Útil para auditoría y validación
- Permite análisis granular
- Incluye todos los campos originales

### Hoja 3: "Solo en [Periodo 1]"
Registros que **NO** aparecen en el Periodo 2:
- 🚨 Clientes que dejaron de comprar
- 🚨 Productos descontinuados
- 🚨 Servicios no renovados
- 💡 Oportunidades de recuperación

### Hoja 4: "Solo en [Periodo 2]"
Registros que **NO** aparecen en el Periodo 1:
- 🎉 Nuevos clientes captados
- 🎉 Nuevos productos vendidos
- 🎉 Expansión de servicios
- 📈 Éxitos comerciales

### Hoja 5: "Comunes en ambos"
Registros que aparecen en **ambos** periodos:
- ✅ Clientes recurrentes
- ✅ Productos con demanda sostenida
- 📊 Análisis de crecimiento/decrecimiento
- 💎 Base de clientes fidelizados

## 🎯 Casos de Uso

### 1. Comparación Trimestral
```
Periodo 1: Q1 2024 (Enero-Marzo 2024)
Periodo 2: Q1 2023 (Enero-Marzo 2023)
Objetivo: Medir crecimiento interanual
```

### 2. Comparación Mensual
```
Periodo 1: Enero 2024
Periodo 2: Febrero 2024
Objetivo: Evolución mes a mes
```

### 3. Análisis de Campañas
```
Periodo 1: Pre-campaña (Noviembre 2023)
Periodo 2: Durante campaña (Diciembre 2023)
Objetivo: Medir impacto de acciones comerciales
```

### 4. Evaluación de Representantes
```
Filtrar por: SalesRepresentative = "Juan Pérez"
Periodo 1: H1 2023
Periodo 2: H1 2024
Objetivo: Evaluar performance individual
```

## 📈 Interpretación de Resultados

### Crecimiento Positivo ✅
```
Cliente A, Producto X:
- Cantidad P1: 100
- Cantidad P2: 150
- Diferencia: +50 (↑ 50%)
→ Incremento en ventas, posible oportunidad de expansión
```

### Decrecimiento ⚠️
```
Cliente B, Producto Y:
- Importe P1: €10,000
- Importe P2: €7,000
- Diferencia: -€3,000 (↓ 30%)
→ Reducción de ventas, requiere acción comercial
```

### Cliente Perdido 🚨
```
Cliente C, Producto Z:
- P1: Presente (€5,000)
- P2: Ausente (€0)
- Ubicación: Hoja "Solo en Periodo 1"
→ Cliente no renovó, contactar para recuperación
```

### Nueva Venta 🎉
```
Cliente D, Producto W:
- P1: Ausente (€0)
- P2: Presente (€8,000)
- Ubicación: Hoja "Solo en Periodo 2"
→ Nueva captación exitosa
```

## 📊 KPIs Principales

### Tasa de Retención
```
Tasa = (Registros comunes / Total registros P1) × 100
Objetivo: > 80%
```

### Tasa de Captación
```
Tasa = (Registros solo P2 / Total registros P2) × 100
Objetivo: Depende de estrategia comercial
```

### Crecimiento en Ventas Comunes
```
Crecimiento = ((Importe P2 comunes - Importe P1 comunes) / Importe P1 comunes) × 100
Objetivo: > 5% interanual
```

## 🔧 Solución de Problemas

### Error: "Columna no encontrada"
**Causa:** Nombres de columnas diferentes a los esperados

**Solución:**
1. Usa los selectores de mapeo de columnas en la sección "🛠️ Asignación de columnas"
2. Verifica los nombres exactos en tu exportación de Power BI

### Error: "No se pueden convertir valores a numérico"
**Causa:** Columnas de cantidad o importe contienen texto

**Solución:**
- Verifica que en Power BI los campos sean numéricos
- Elimina filas de totales/subtotales antes de exportar

### Caracteres Especiales (ñ, acentos)
**Causa:** Problemas de codificación en CSV

**Solución:**
- Usa formato Excel (`.xlsx`) en lugar de CSV
- Si usas CSV, asegura codificación UTF-8 con BOM

### Resultados Inesperados en Importes
**Nota:** El campo EUR debe contener el **importe total**, no el precio unitario.

La aplicación usa directamente este valor sin multiplicar por cantidad.

## 💡 Mejores Prácticas

### Nomenclatura de Periodos
- ✅ "Q1 2024", "Q2 2024" (comparaciones trimestrales)
- ✅ "Enero 2024", "Enero 2023" (comparaciones mensuales)
- ✅ "H1 2024", "H2 2024" (semestres)
- ❌ "Periodo 1", "Periodo 2" (ambiguo, dificulta seguimiento)

### Frecuencia de Análisis Recomendada
- **Mensual:** Para seguimiento continuo y acciones rápidas
- **Trimestral:** Para reporting ejecutivo y revisión de estrategia
- **Anual:** Para planificación estratégica a largo plazo

### Almacenamiento de Resultados
Guarda los archivos con fecha para mantener historial:
```
comparativa_Q1_2024_vs_Q1_2023_20241030.xlsx
```

## ⚙️ Consideraciones Técnicas

### Rendimiento
- **Límite recomendado:** 100,000 registros por archivo
- **Tiempo de procesamiento:** 5-30 segundos según tamaño
- **Memoria:** Suficiente para archivos típicos de Power BI

### Limitaciones
- Solo soporta EUR como moneda
- Requiere nombres de columnas consistentes
- No consolida automáticamente clientes con nombres ligeramente diferentes

## 🤝 Contribuciones

Si encuentras errores o tienes sugerencias de mejora:
1. Abre un issue en GitHub
2. Describe el problema o mejora
3. Incluye ejemplos si es posible

## 📝 Licencia

Este proyecto es de uso interno para análisis de ventas.

## 👥 Soporte

Para preguntas o soporte:
- Consulta esta documentación
- Revisa la sección de "Solución de Problemas"
- Contacta al equipo de datos

---

**Versión:** 1.0  
**Última actualización:** Octubre 2024  
**Tecnología:** Streamlit + Pandas + OpenPyXL
