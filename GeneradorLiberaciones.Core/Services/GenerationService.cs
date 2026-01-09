using GemBox.Document;
using GeneradorLiberaciones.Core.Config;
using GeneradorLiberaciones.Core.Models;
using GeneradorLiberaciones.Core.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


namespace GeneradorLiberaciones.Core.Services;

public class GenerationOptions
{
    public bool GenerarLiberaciones { get; set; } = true;
    public bool GenerarCalificaciones { get; set; } = true;

    public string TemplateLiberacionPath { get; set; } = "Templates/Liberacion.docx";
    public string TemplateCalificacionPath { get; set; } = "Templates/Calificacion.docx";

    public string OutputRoot { get; set; } = "Salida";

    // Folios iniciales
    public int FolioInicialLiberacion { get; set; } = 1;
    public int FolioInicialCalificacion { get; set; } = 1;
}

public class GenerationResult
{
    public int DocumentosGenerados { get; set; }
    public int? LastFolioLiberacion { get; set; }
    public int? LastFolioCalificacion { get; set; }
    public HashSet<string> NoControlEncontradosEnExcel { get; } =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase);
}

public class GenerationService
{
    private readonly ExcelReader excelReader;
    private readonly TemplateEngine templateEngine;
    private static string? cachedLibreOfficePath;

    // Regex compiladas para un rendimiento óptimo en validaciones repetitivas.
    private static readonly Regex digitsOnlyRegex = new Regex("^[0-9]{8}$", RegexOptions.Compiled);
    private static readonly Regex letterAndDigitsRegex = new Regex("^[A-Z][0-9]{8}$", RegexOptions.Compiled);

    public GenerationService(ExcelReader excelReader, TemplateEngine templateEngine)
    {
        this.excelReader = excelReader;
        this.templateEngine = templateEngine;
    }

    /// <summary>
    /// - rutaExcel: ruta del archivo .xlsx
    /// - options: configuración de generación
    /// - noControlYaGenerados: NO_CONTROL a excluir (ya tienen documentos)
    /// - excluirYaGenerados: si true, se omiten los anteriores
    /// </summary>
    public GenerationResult GenerarTodo(
    string rutaExcel,
    GenerationOptions options,
    bool soloListado,
    HashSet<string>? noControlObjetivo,
    HashSet<string>? noControlYaGenerados = null,
    bool excluirYaGenerados = false)
    {
        if (!Directory.Exists(options.OutputRoot))
        {
            Directory.CreateDirectory(options.OutputRoot);
        }

        var registros = excelReader.Leer(rutaExcel);
        Console.WriteLine($"[INFO] Registros leídos desde el Excel: {registros.Count}");

        int documentosGenerados = 0;

        int folioLiberacionActual = options.FolioInicialLiberacion;
        int folioCalificacionActual = options.FolioInicialCalificacion;

        int? lastFolioLiberacionUsado = null;
        int? lastFolioCalificacionUsado = null;

        var yaGenerados = noControlYaGenerados != null
            ? new HashSet<string>(noControlYaGenerados, StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var generadosEstaVez = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var errores = new List<string>();

        // Resultado a devolver (para poder ir llenando NoControlEncontradosEnExcel)
        var result = new GenerationResult();

        foreach (var record in registros)
        {
            var noControlKey = (record.NoControl ?? string.Empty).Trim();
            var noControlDisplay = string.IsNullOrWhiteSpace(noControlKey) ? "SIN_NO_CONTROL" : noControlKey;
            var nombreDisplay = string.IsNullOrWhiteSpace(record.Nombre) ? "SIN_NOMBRE" : record.Nombre.Trim();

            // Registrar que ESTE NO_CONTROL existe en el Excel
            if (!string.IsNullOrEmpty(noControlKey))
            {
                result.NoControlEncontradosEnExcel.Add(noControlKey);
            }

            // 0) Modo "soloListado": solo generamos para NO_CONTROL que estén en la lista objetivo
            if (soloListado)
            {
                // Si no hay listado, o este NO_CONTROL no está en él, se omite
                if (string.IsNullOrEmpty(noControlKey) ||
                    noControlObjetivo == null ||
                    !noControlObjetivo.Contains(noControlKey))
                {
                    // No es error; simplemente no está en la lista a procesar
                    continue;
                }
            }

            // 1) Excluir ya generados (modo opción 2)
            if (excluirYaGenerados &&
                !string.IsNullOrEmpty(noControlKey) &&
                yaGenerados.Contains(noControlKey))
            {
                Console.WriteLine($"[SKIP-EXISTE] {noControlDisplay} - {nombreDisplay} - ya tenía documentos generados, se omite.");
                continue;
            }

            string? errorMessage = null;

            // 2) Validaciones fuertes → se loguean en archivo de errores

            // Opción 3: No. de control inválido
            if (!EsNoControlValido(noControlKey))
            {
                errorMessage = "Errror 3";
            }
            // Opción 2: Fechas nulas / inválidas
            else if (record.Inicio == null || record.Termino == null)
            {
                errorMessage = "Errror 4";
            }
            // Opción 1: Periodo distinto de 6 meses
            else if (!EsPeriodoDeSeisMeses(record.Inicio.Value, record.Termino.Value))
            {
                errorMessage = "Errror 1";
            }
            // Opción 4: Programa no permitido
            else if (!EsProgramaValido(record.Programa, out var programaNormalizado))
            {
                errorMessage = "Errror 2";
            }
            else
            {
                // Programa normalizado para sacar siempre el mismo texto
                record.Programa = programaNormalizado!;

                // 3) Condiciones de datos incompletos (helper externo)
                // Estos NO van al log de errores, solo consola.
                if (!GenerateConditionHelper.DebeGenerarDocumentos(record))
                {
                    Console.WriteLine($"[SKIP] {noControlDisplay} - {nombreDisplay} - datos incompletos, no se genera documento.");
                    continue;
                }
            }

            if (!string.IsNullOrEmpty(errorMessage))
            {
                var lineaError = $"{noControlDisplay} - {nombreDisplay} - {errorMessage}";
                Console.WriteLine($"[ERROR] {lineaError}");
                errores.Add(lineaError);
                continue;
            }

            // 4) Pasa todos los filtros → generar documentos
            var carreraFolder = SanitizeFolderName(record.Carrera);
            bool generoAlgoParaEsteAlumno = false;

            if (options.GenerarLiberaciones)
            {
                record.FolioLiberacion = folioLiberacionActual.ToString();
                lastFolioLiberacionUsado = folioLiberacionActual;
                folioLiberacionActual++;

                var baseName = $"Liberacion_{SanitizeFileName(record.Nombre)}_{record.NoControl}";
                var carpetaLib = Path.Combine(options.OutputRoot, "Liberaciones", carreraFolder);

                if (!Directory.Exists(carpetaLib))
                    Directory.CreateDirectory(carpetaLib);

                var docxPath = Path.Combine(carpetaLib, baseName + ".docx");

                // 1) Generar DOCX desde la plantilla
                templateEngine.GenerarDocumento(
                    options.TemplateLiberacionPath,
                    docxPath,
                    record,
                    PlaceholderConfigs.Liberacion);

                // 2) Convertir a PDF usando LibreOffice
                ConvertirDocxAPdfConLibreOffice(docxPath);

                documentosGenerados++;
                generoAlgoParaEsteAlumno = true;
            }


            if (options.GenerarCalificaciones)
            {
                record.FolioCalificacion = folioCalificacionActual.ToString();
                lastFolioCalificacionUsado = folioCalificacionActual;
                folioCalificacionActual++;

                var baseName = $"Calificacion_{SanitizeFileName(record.Nombre)}_{record.NoControl}";
                var carpetaCal = Path.Combine(options.OutputRoot, "Calificaciones", carreraFolder);

                if (!Directory.Exists(carpetaCal))
                    Directory.CreateDirectory(carpetaCal);

                var docxPath = Path.Combine(carpetaCal, baseName + ".docx");

                // 1) Generar DOCX desde la plantilla
                templateEngine.GenerarDocumento(
                    options.TemplateCalificacionPath,
                    docxPath,
                    record,
                    PlaceholderConfigs.Calificacion);

                // 2) Convertir a PDF usando LibreOffice
                ConvertirDocxAPdfConLibreOffice(docxPath);

                documentosGenerados++;
                generoAlgoParaEsteAlumno = true;
            }


            if (generoAlgoParaEsteAlumno && !string.IsNullOrEmpty(noControlKey))
            {
                generadosEstaVez.Add(noControlKey);
            }
        }

        Console.WriteLine($"[INFO] Documentos generados en esta ejecución: {documentosGenerados}");

        // Unión de previos + nuevos
        yaGenerados.UnionWith(generadosEstaVez);

        // Log acumulado de NO_CONTROL SIEMPRE en OutputRoot con nombre nuevo
        GenerarLogNoControl(options.OutputRoot, yaGenerados);

        // Log de errores
        GenerarLogErrores(options.OutputRoot, errores);

        // Rellenar result y devolver
        result.DocumentosGenerados = documentosGenerados;
        result.LastFolioLiberacion = lastFolioLiberacionUsado;
        result.LastFolioCalificacion = lastFolioCalificacionUsado;

        return result;
    }

    private static string? GetLibreOfficePath()
    {
        // 1) Variable de entorno opcional (por si algún día lo necesitas)
        var fromEnv = Environment.GetEnvironmentVariable("LIBREOFFICE_PATH");
        if (!string.IsNullOrWhiteSpace(fromEnv) && File.Exists(fromEnv))
        {
            cachedLibreOfficePath = fromEnv;
            return cachedLibreOfficePath;
        }

        // 2) Rutas típicas de instalación en Windows
        string[] posibles =
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        };

        foreach (var p in posibles)
        {
            if (File.Exists(p))
            {
                cachedLibreOfficePath = p;
                return cachedLibreOfficePath;
            }
        }

        // 3) Si no se encontró, devolvemos null
        return null;
    }

    private void ConvertirDocxAPdfConLibreOffice(string docxPath)
    {
        var libreOfficePath = GetLibreOfficePath();
        if (libreOfficePath == null)
        {
            // Aquí lanzamos una excepción clara que luego capturas en la UI
            throw new InvalidOperationException(
                "No se encontró LibreOffice en esta computadora.\n\n" +
                "Instala LibreOffice para poder generar los PDF.\n" +
                "Si ya lo tienes instalado en una ruta diferente, " +
                "puedes crear una variable de entorno 'LIBREOFFICE_PATH' " +
                "apuntando al archivo 'soffice.exe'.");
        }

        var outputDir = Path.GetDirectoryName(docxPath) ?? ".";

        var startInfo = new ProcessStartInfo
        {
            FileName = libreOfficePath,
            Arguments = $"--headless --convert-to pdf \"{docxPath}\" --outdir \"{outputDir}\"",
            CreateNoWindow = true,
            UseShellExecute = false
        };

        using var process = Process.Start(startInfo);
        process!.WaitForExit();

        // Opcional: borrar el DOCX después de convertir
        if (File.Exists(docxPath))
             File.Delete(docxPath);
    }
    // === Validaciones auxiliares ===

    private static bool     EsNoControlValido(string? noControl)
    {
        if (string.IsNullOrWhiteSpace(noControl))
            return false;

        noControl = noControl.Trim().ToUpperInvariant();

        return digitsOnlyRegex.IsMatch(noControl) || letterAndDigitsRegex.IsMatch(noControl);
    }

    private static bool EsPeriodoDeSeisMeses(DateTime inicio, DateTime termino)
    {
        return inicio.Date.AddMonths(6) == termino.Date;
    }

    /// Programa válido: SOPORTE TÉCNICO, ASISTENCIA TÉCNICA, INNOVATEC, ASESORÍAS (con o sin acento)
    private static bool EsProgramaValido(string? programa, out string? programaNormalizado)
    {
        programaNormalizado = null;

        if (string.IsNullOrWhiteSpace(programa))
            return false;

        var raw = programa.Trim().ToUpperInvariant();
        var sinAcentos = RemoveDiacritics(raw);

        switch (sinAcentos)
        {
            case "SOPORTE TECNICO":
                programaNormalizado = "SOPORTE TÉCNICO";
                return true;
            case "ASISTENCIA TECNICA":
                programaNormalizado = "ASISTENCIA TÉCNICA";
                return true;
            case "INNOVATEC":
                programaNormalizado = "INNOVATEC";
                return true;
            case "ASESORIAS":
                programaNormalizado = "ASESORÍAS";
                return true;
            default:
                return false;
        }
    }

    private static string RemoveDiacritics(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var normalized = text.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();

        foreach (var c in normalized)
        {
            var unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(c);
            }
        }

        return sb.ToString().Normalize(NormalizationForm.FormC);
    }

    // === Logs ===
    private static void GenerarLogNoControl(string outputRoot, HashSet<string> todosGenerados)
    {
        try
        {
            if (!Directory.Exists(outputRoot))
                Directory.CreateDirectory(outputRoot);

            var logName = $"NoControlGenerados_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            var logPath = Path.Combine(outputRoot, logName);

            var lineas = new List<string> { "NO_CONTROL" };
            lineas.AddRange(todosGenerados.OrderBy(x => x));

            File.WriteAllLines(logPath, lineas);

            Console.WriteLine($"[INFO] Log de generados: {logPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("[WARN] No se pudo generar el archivo de log de NO_CONTROL:");
            Console.WriteLine(ex.Message);
        }
    }

    private static void GenerarLogErrores(string outputRoot, List<string> errores)
    {
        try
        {
            if (errores.Count == 0)
            {
                Console.WriteLine("[INFO] No hubo errores de formato (no. control / fechas / programa).");
                return;
            }

            if (!Directory.Exists(outputRoot))
                Directory.CreateDirectory(outputRoot);

            var logName = $"ErroresGeneracion_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            var logPath = Path.Combine(outputRoot, logName);

            var lineas = new List<string>
            {
                "Listado de registros que NO generaron documento por errores de formato:",
                "  - Las fechas de inicio y término no forman un periodo exacto de 6 meses.",
                "  - Opción 2: Fecha de inicio o término nula o con formato inválido.",
                "  - Opción 3: Número de control inválido (debe ser 8 dígitos o una letra seguida de 8 dígitos).",
                "  - Opción 4: nombre de programa no existente o no permitido.",
                ""
            };

            lineas.AddRange(errores);

            File.WriteAllLines(logPath, lineas);

            Console.WriteLine($"[INFO] Log de errores: {logPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("[WARN] No se pudo generar el archivo de log de errores:");
            Console.WriteLine(ex.Message);
        }
    }

    // === Utilidades de nombres de archivo/carpeta ===

    private static string SanitizeFileName(string input)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var safe = new string((input ?? string.Empty).Where(c => !invalid.Contains(c)).ToArray());
        return string.IsNullOrWhiteSpace(safe) ? "SIN_NOMBRE" : safe;
    }

    private static string SanitizeFolderName(string input)
    {
        var invalid = Path.GetInvalidPathChars().Concat(Path.GetInvalidFileNameChars()).ToArray();
        var safe = new string((input ?? string.Empty).Where(c => !invalid.Contains(c)).ToArray());
        return string.IsNullOrWhiteSpace(safe) ? "SIN_CARRERA" : safe;
    }
}
