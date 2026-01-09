using System;
using System.Collections.Generic;
using System.Text;

namespace GeneradorLiberaciones.Core.Utils;

public static class DateUtils
{
    private static readonly string[] Meses =
    {
        "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    };

    public static string? FormatearFechaLarga(DateTime? fecha)
    {
        if (fecha == null) return null;
        var f = fecha.Value;
        var mes = Meses[f.Month];
        return $"{f.Day} de {mes} del {f.Year}";
    }
}