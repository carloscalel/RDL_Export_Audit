using Microsoft.Data.SqlClient;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.IO.Compression;
using System.Text.Json;
using ClosedXML.Excel;
using System.Data;

class Program
{
    static void Main()
    {
        Console.WriteLine("Ingreso de datos de conexión SQL Server Reporting Services:");

        Console.Write("Servidor SQL (ej. localhost): ");
        string servidor = Console.ReadLine()?.Trim() ?? "localhost";

        Console.Write("Base de datos (por defecto: ReportServer): ");
        string? baseDatos = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(baseDatos)) baseDatos = "ReportServer";

        Console.Write("Ruta de carpeta de salida (ej. C:\\ExportSSRS): ");
        string carpetaSalida = Console.ReadLine()?.Trim() ?? @"C:\ExportSSRS";

        if (!Directory.Exists(carpetaSalida))
            Directory.CreateDirectory(carpetaSalida);

        string rutaExcel = Path.Combine(carpetaSalida, "Auditoria_RDL_SQLServer2022.xlsx");

        // FILTRO OPCIONAL DE REPORTES
        Console.Write("Ingrese nombres de reportes separados por coma (Enter = todos): ");
        string filtroInput = Console.ReadLine() ?? "";

        var reportesFiltrar = filtroInput
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        // EXPORTAR REPORTES RDL
        Console.Write("¿Desea exportar también los archivos RDL? (S/N): ");
        bool exportarRdl = (Console.ReadLine()?.Trim().ToUpper() == "S");

        string carpetaRdl = Path.Combine(carpetaSalida, "RDL_Exportados");

        if (exportarRdl && !Directory.Exists(carpetaRdl))
            Directory.CreateDirectory(carpetaRdl);

        // CARGAR REGLAS
        var motorReglas = new MotorReglas("reglas.json");

        var listaHallazgos = new List<HallazgoExcel>();

        string connectionString =
            $"Server={servidor};Database={baseDatos};Integrated Security=true;Encrypt=True;TrustServerCertificate=True;";

        string queryBase = @"
            SELECT [Name], [Path], [Content]
            FROM [dbo].[Catalog]
            WHERE [Type] = 2";

        string query = queryBase;

        if (reportesFiltrar.Any())
        {
            var parametros = reportesFiltrar
                .Select((r, i) => $"@rep{i}")
                .ToList();

            query += $" AND [Name] IN ({string.Join(",", parametros)})";
        }

        int contador = 0;

        using (SqlConnection conexion = new SqlConnection(connectionString))
        {
            conexion.Open();

            using (SqlCommand comando = new SqlCommand(query, conexion))
            {
                if (reportesFiltrar.Any())
                {
                    for (int i = 0; i < reportesFiltrar.Count; i++)
                    {
                        comando.Parameters.AddWithValue($"@rep{i}", reportesFiltrar[i]);
                    }
                }

                using (SqlDataReader lector = comando.ExecuteReader())
                {
                    while (lector.Read())
                    {
                        string nombre = lector.GetString(0);
                        string path = lector.GetString(1);
                        byte[] contenido = (byte[])lector["Content"];

                        string xmlString;

                        if (contenido.Length > 2 && contenido[0] == 0x1F && contenido[1] == 0x8B)
                        {
                            using var ms = new MemoryStream(contenido);
                            using var gzip = new GZipStream(ms, CompressionMode.Decompress);
                            using var reader = new StreamReader(gzip, Encoding.UTF8);
                            xmlString = reader.ReadToEnd();
                        }
                        else
                        {
                            xmlString = Encoding.UTF8.GetString(contenido);
                        }

                        int inicioXml = xmlString.IndexOf('<');
                        if (inicioXml > 0)
                            xmlString = xmlString.Substring(inicioXml);

                        if (!xmlString.TrimStart().StartsWith("<"))
                            continue;

                        // EXPORTAR RDL SI SE SOLICITÓ
                        if (exportarRdl)
                        {
                            string rutaRelativa = path.Replace("/", "\\").TrimStart('\\');
                            string rutaCompleta = Path.Combine(carpetaRdl, rutaRelativa + ".rdl");

                            string? directorio = Path.GetDirectoryName(rutaCompleta);
                            if (!Directory.Exists(directorio))
                                Directory.CreateDirectory(directorio!);

                            File.WriteAllText(rutaCompleta, xmlString, Encoding.UTF8);
                        }

                        XDocument doc = XDocument.Parse(xmlString);

                        var dataSets = doc.Descendants()
                            .Where(x => x.Name.LocalName == "DataSet");

                        foreach (var ds in dataSets)
                        {
                            string dsName = ds.Attribute("Name")?.Value ?? "SinNombre";

                            var commandTextNode = ds.Descendants()
                                .FirstOrDefault(x => x.Name.LocalName == "CommandText");

                            if (commandTextNode == null)
                                continue;

                            string querySql = commandTextNode.Value;

                            int score = 0;
                            List<string> hallazgos = new List<string>();

                            // VALIDACIÓN POR REGLAS JSON (SQL + XML)
                            //foreach (var regla in motorReglas.Reglas.Values)
                            //{
                            //    if (string.IsNullOrWhiteSpace(regla.Patron))
                            //        continue;

                            //    bool matchSql = Regex.IsMatch(querySql ?? "", regla.Patron, RegexOptions.IgnoreCase);
                            //    bool matchXml = Regex.IsMatch(xmlString ?? "", regla.Patron, RegexOptions.IgnoreCase);

                            //    if (matchSql || matchXml)
                            //    {
                            //        score += regla.Puntaje;
                            //        hallazgos.Add(regla.Codigo ?? "REGLA_SIN_CODIGO");
                            //    }
                            //}
                            foreach (var regla in motorReglas.Reglas.Values)
                            {
                                if (string.IsNullOrWhiteSpace(regla.Patron))
                                    continue;

                                bool match = false;

                                if (regla.Ambito == "SQL")
                                {
                                    match = Regex.IsMatch(querySql ?? "", regla.Patron, RegexOptions.IgnoreCase);
                                }
                                else if (regla.Ambito == "XML")
                                {
                                    match = Regex.IsMatch(xmlString ?? "", regla.Patron, RegexOptions.IgnoreCase);
                                }
                                else // BOTH o null
                                {
                                    match =
                                        Regex.IsMatch(querySql ?? "", regla.Patron, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(xmlString ?? "", regla.Patron, RegexOptions.IgnoreCase);
                                }

                                if (match)
                                {
                                    score += regla.Puntaje;
                                    hallazgos.Add(regla.Codigo ?? "REGLA_SIN_CODIGO");
                                }
                            }

                            // VALIDACIÓN DINÁMICA DEL RESULTSET
                            try
                            {
                                string connMaster =
                                    $"Server={servidor};Database=master;Integrated Security=true;Encrypt=True;TrustServerCertificate=True;";

                                using SqlConnection conexionDescribe = new SqlConnection(connMaster);
                                conexionDescribe.Open();

                                using SqlCommand cmdDescribe =
                                    new SqlCommand("sp_describe_first_result_set", conexionDescribe);

                                cmdDescribe.CommandType = CommandType.StoredProcedure;

                                string sqlSeguro = querySql?.Trim() ?? string.Empty;
                                sqlSeguro = sqlSeguro.Replace("\0", "");

                                SqlParameter pTsql = new SqlParameter("@tsql", SqlDbType.NVarChar);
                                pTsql.Size = -1;
                                pTsql.Value = sqlSeguro;

                                cmdDescribe.Parameters.Add(pTsql);
                                cmdDescribe.Parameters.Add(new SqlParameter("@params", SqlDbType.NVarChar) { Size = -1, Value = DBNull.Value });
                                cmdDescribe.Parameters.Add(new SqlParameter("@browse_information_mode", SqlDbType.Int) { Value = 0 });

                                using SqlDataReader dr = cmdDescribe.ExecuteReader();

                                while (dr.Read())
                                {
                                    string? tipo = dr["system_type_name"]?.ToString();

                                    if (!string.IsNullOrEmpty(tipo) &&
                                        Regex.IsMatch(tipo, @"text|ntext|image", RegexOptions.IgnoreCase))
                                    {
                                        score += 35;
                                        hallazgos.Add("RESULTSET_DEPRECATED");
                                    }
                                }
                            }
                            catch
                            {
                                hallazgos.Add("NO_SE_PUDO_ANALIZAR_RESULTSET");
                            }

                            string severidad = score switch
                            {
                                <= 20 => "Bajo",
                                <= 50 => "Medio",
                                <= 80 => "Alto",
                                _ => "Critico"
                            };

                            listaHallazgos.Add(new HallazgoExcel
                            {
                                Reporte = nombre,
                                Path = path,
                                DataSet = dsName,
                                Score = score,
                                Severidad = severidad,
                                Detalle = hallazgos.Any()
                                    ? string.Join(", ", hallazgos)
                                    : "Sin hallazgos"
                            });
                        }

                        contador++;
                        Console.WriteLine($"Analizado: {nombre}");
                    }
                }
            }
        }

        ExportadorExcelAuditoria.GenerarExcel(rutaExcel, listaHallazgos, motorReglas);

        Console.WriteLine($"\nExportación y auditoría completada. Total reportes: {contador}");
        Console.WriteLine($"Excel generado en: {rutaExcel}");
        Console.ReadKey();
    }
}

// ============================
// CLASE REGLAS
// ============================

public class ReglaAuditoria
{
    public string? Codigo { get; set; }
    public string? Descripcion { get; set; }
    public int Puntaje { get; set; }
    public string? Severidad { get; set; }
    public string? Patron { get; set; }
    public string? Ambito { get; set; } // SQL | XML | BOTH
}

public class MotorReglas
{
    public Dictionary<string, ReglaAuditoria> Reglas { get; private set; }

    public MotorReglas(string rutaJson)
    {
        var json = File.ReadAllText(rutaJson);
        var lista = JsonSerializer.Deserialize<List<ReglaAuditoria>>(json);
        Reglas = lista?.ToDictionary(r => r.Codigo ?? Guid.NewGuid().ToString(), r => r)
                 ?? new Dictionary<string, ReglaAuditoria>();
    }
}

// ============================
// CLASE EXCEL
// ============================

public class HallazgoExcel
{
    public string? Reporte { get; set; }
    public string? Path { get; set; }
    public string? DataSet { get; set; }
    public int Score { get; set; }
    public string? Severidad { get; set; }
    public string? Detalle { get; set; }
}

public class ExportadorExcelAuditoria
{
    public static void GenerarExcel(
        string rutaExcel,
        List<HallazgoExcel> hallazgos,
        MotorReglas motorReglas)
    {
        using var wb = new XLWorkbook();

        var wsHallazgos = wb.Worksheets.Add("Hallazgos");

        wsHallazgos.Cell(1, 1).Value = "Reporte";
        wsHallazgos.Cell(1, 2).Value = "Path";
        wsHallazgos.Cell(1, 3).Value = "DataSet";
        wsHallazgos.Cell(1, 4).Value = "Score";
        wsHallazgos.Cell(1, 5).Value = "Severidad";
        wsHallazgos.Cell(1, 6).Value = "Detalle";

        int fila = 2;

        foreach (var h in hallazgos)
        {
            wsHallazgos.Cell(fila, 1).Value = h.Reporte;
            wsHallazgos.Cell(fila, 2).Value = h.Path;
            wsHallazgos.Cell(fila, 3).Value = h.DataSet;
            wsHallazgos.Cell(fila, 4).Value = h.Score;
            wsHallazgos.Cell(fila, 5).Value = h.Severidad;
            wsHallazgos.Cell(fila, 6).Value = h.Detalle;
            fila++;
        }

        wsHallazgos.Columns().AdjustToContents();

        var wsCriterios = wb.Worksheets.Add("Matriz_Criterios");

        wsCriterios.Cell(1, 1).Value = "Codigo";
        wsCriterios.Cell(1, 2).Value = "Descripcion";
        wsCriterios.Cell(1, 3).Value = "Puntaje";
        wsCriterios.Cell(1, 4).Value = "Severidad";
        wsCriterios.Cell(1, 5).Value = "Patron Regex";

        int fila2 = 2;

        foreach (var regla in motorReglas.Reglas.Values)
        {
            wsCriterios.Cell(fila2, 1).Value = regla.Codigo;
            wsCriterios.Cell(fila2, 2).Value = regla.Descripcion;
            wsCriterios.Cell(fila2, 3).Value = regla.Puntaje;
            wsCriterios.Cell(fila2, 4).Value = regla.Severidad;
            wsCriterios.Cell(fila2, 5).Value = regla.Patron;
            fila2++;
        }

        wsCriterios.Columns().AdjustToContents();

        wb.SaveAs(rutaExcel);
    }
}