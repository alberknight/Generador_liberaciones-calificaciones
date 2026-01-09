using System;
using System.Collections.Generic;
using System.Text;

namespace GeneradorLiberaciones.Core.Utils;

public static class CarreraHelper
{
    public static string DeducirCarrera(string? nombreHoja)
    {
        if (string.IsNullOrWhiteSpace(nombreHoja))
            return "CARRERA DESCONOCIDA";

        var key = nombreHoja.Trim().ToUpperInvariant();

        // Por si hay errores de captura tipo "MECAINCA"
        if (key.Contains("MECAINCA"))
            key = "MECANICA";

        // Si quieres, aquí también podrías normalizar acentos:
        // key = key.Replace("Á", "A").Replace("É","E")...

        return key switch
        {
            // 1) INDUSTRIAL
            "INDUSTRIAL" =>
                "INDUSTRIAL",

            // 2) MECÁNICA (cubriendo variantes)
            "MECANICA" or "MECÁNICA" =>
                "MECÁNICA",

            // 3) BIOQUÍMICA
            "BIOQUIMICA" or "BIOQUÍMICA" =>
                "BIOQUÍMICA",

            // 4) GESTIÓN
            "GESTION" or "GESTIÓN" =>
                "EN GESTIÓN EMPRESARIAL",

            // 5) SISTEMAS
            "SISTEMAS" or "SISTEMAS COMPUTACIONALES" or "ISC" =>
                "EN SISTEMAS COMPUTACIONALES",

            // 6) QUÍMICA
            "QUIMICA" or "QUÍMICA" =>
                "QUÍMICA",

            // 7) ELÉCTRICA
            "ELECTRICA" or "ELÉCTRICA" =>
                "ELÉCTRICA",

            // 8) LOGÍSTICA
            "LOGISTICA" or "LOGÍSTICA" =>
                "EN LOGÍSTICA",

            // 9) ELECTRÓNICA
            "ELECTRONICA" or "ELECTRÓNICA" =>
                "ELECTRÓNICA",

            // Si llega algo raro, usamos el nombre tal cual
            _ => nombreHoja
        };
    }
}