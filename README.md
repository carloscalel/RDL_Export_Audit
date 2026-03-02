# Auditoría Técnica de Reportes SSRS (SQL Server 2022)

Herramienta de consola en .NET para auditar reportes SSRS

Permite:

- Analizar queries dentro de RDL
- Detectar malas prácticas SQL mediante reglas configurables (JSON)
- Evaluar tipos de datos del resultset usando sp_describe_first_result_set
- Generar reporte Excel con hallazgos
- Exportar archivos RDL físicamente
- Filtrar reportes específicos
- Calcular score y severidad automática

---

## 🚀 Características

✔ Motor de reglas configurable vía JSON  
✔ Análisis SQL y XML  
✔ Validación dinámica del resultset  
✔ Exportación Excel (Hallazgos + Matriz de Criterios)  
✔ Exportación opcional de archivos .rdl  
✔ Filtro por nombre de reporte  
✔ Clasificación por severidad (Bajo / Medio / Alto / Crítico)  

---

## 🧠 Arquitectura

- Fuente de datos: ReportServer.dbo.Catalog
- Tipo 2 = Reportes RDL
- Reglas definidas en `reglas.json`
- Evaluación vía Regex
- Exportación usando ClosedXML

---

## 📦 Requisitos

- .NET 6 o superior
- SQL Server 2016+ (para sp_describe_first_result_set)
- Permisos de lectura en ReportServer
- Permisos de escritura en carpeta destino

---

## 📥 Instalación de Paquetes NuGet

Instalar los siguientes paquetes:

### 1️⃣ Microsoft.Data.SqlClient

Proveedor oficial para SQL Server.

### 2️⃣ ClosedXML

Para generación de archivos Excel (.xlsx).

### 3️⃣ ClosedXML instala automáticamente:

DocumentFormat.OpenXml
using ClosedXML.Excel;

