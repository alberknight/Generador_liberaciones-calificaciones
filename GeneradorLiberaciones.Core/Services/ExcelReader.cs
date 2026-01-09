using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using GeneradorLiberaciones.Core.Models;
using GeneradorLiberaciones.Core.Utils;

namespace GeneradorLiberaciones.Core.Services;

public class ExcelReader
{
    public List<StudentRecord> Leer(string rutaExcel)
    {
        var lista = new List<StudentRecord>();

        using var wb = new XLWorkbook(rutaExcel);
        foreach (var ws in wb.Worksheets)
        {
            var carrera = CarreraHelper.DeducirCarrera(ws.Name);

            var firstRow = ws.FirstRowUsed();
            if (firstRow == null) continue;

            int headerRowNumber = firstRow.RowNumber();
            int lastRowNumber = ws.LastRowUsed().RowNumber();

            var headerCells = ws.Row(headerRowNumber).CellsUsed();
            var columnas = headerCells.ToDictionary(
                c => c.GetString().Trim().ToUpperInvariant(),
                c => c.Address.ColumnNumber);

            for (int rowNum = headerRowNumber + 1; rowNum <= lastRowNumber; rowNum++)
            {
                var row = ws.Row(rowNum);

                string nombre = GetString(row, columnas, "NOMBRE");
                if (string.IsNullOrWhiteSpace(nombre))
                    continue;

                var record = new StudentRecord
                {
                    Nombre = nombre,
                    NoControl = GetString(row, columnas, "NO. CONTROL"),
                    Dependencia = GetString(row, columnas, "DEPENDENCIA"),
                    Programa = GetString(row, columnas, "PROGRAMA"),
                    Carrera = carrera,
                    Inicio = GetDate(row, columnas, "FECHA DE INICIO")
                    ?? GetDate(row, columnas, "INICIO"),
                    Termino = GetDate(row, columnas, "FECHA DE TERMINO")
                    ?? GetDate(row, columnas, "TERMINO"),
                    Calificacion = GetAverageCalificacion(row, columnas),
                    FolioLiberacion = GetString(row, columnas, "FOLIO DE LIBERACION"),

                    CartaPresentacion = GetString(row, columnas, "CARTA DE PRESENTACION"),
                    CartaAceptacion = GetString(row, columnas, "CARTA DE ACEPTACION"),
                    CartaSolicitud = GetString(row, columnas, "CARTA SOLICITUD"),
                    CartaCompromiso = GetString(row, columnas, "CARTA COMPROMISO"),
                    Reporte1 = GetString(row, columnas, "REPORTE 1"),
                    Reporte2 = GetString(row, columnas, "REPORTE 2"),
                    Reporte3 = GetString(row, columnas, "REPORTE 3"),
                    CartaTerminacion = GetString(row, columnas, "CARTA DE TERMINACION"),
                    TrabajoFinal = GetString(row, columnas, "TRABAJO FINAL"),
                };

                lista.Add(record);
            }
        }

        return lista;
    }

    private static decimal? GetAverageCalificacion(IXLRow row, Dictionary<string, int> cols)
    {
        var valores = new List<decimal>();

        // Intentamos leer CALIFICACION 1, 2 y 3 (si existen)
        var c1 = GetDecimal(row, cols, "CALIFICACION 1");
        var c2 = GetDecimal(row, cols, "CALIFICACION 2");
        var c3 = GetDecimal(row, cols, "CALIFICACION 3");

        if (c1.HasValue) valores.Add(c1.Value);
        if (c2.HasValue) valores.Add(c2.Value);
        if (c3.HasValue) valores.Add(c3.Value);

        if (valores.Count == 0)
        {
            // Compatibilidad: si no hay 1/2/3, intenta usar CALIFICACION
            return GetDecimal(row, cols, "CALIFICACION");
        }

        // Promedio
        var promedio = valores.Average();

        // Si quieres trabajar siempre con enteros, redondea:
        promedio = Math.Round(promedio, 0, MidpointRounding.AwayFromZero);

        return promedio;
    }


    private static string GetString(IXLRow row, Dictionary<string, int> cols, string colName)
    {
        if (!cols.TryGetValue(colName.ToUpperInvariant(), out var colIndex))
            return string.Empty;

        return row.Cell(colIndex).GetString().Trim();
    }

    private static DateTime? GetDate(IXLRow row, Dictionary<string, int> cols, string colName)
    {
        if (!cols.TryGetValue(colName.ToUpperInvariant(), out var colIndex))
            return null;

        var cell = row.Cell(colIndex);

        if (cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime();

        if (DateTime.TryParse(cell.GetString(), out var dt))
            return dt;

        return null;
    }

    private static decimal? GetDecimal(IXLRow row, Dictionary<string, int> cols, string colName)
    {
        if (!cols.TryGetValue(colName.ToUpperInvariant(), out var colIndex))
            return null;

        var cell = row.Cell(colIndex);

        if (cell.DataType == XLDataType.Number)
            return (decimal)cell.GetDouble();

        if (decimal.TryParse(cell.GetString(), out var dec))
            return dec;

        return null;
    }
}