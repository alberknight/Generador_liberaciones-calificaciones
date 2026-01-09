using System;
using System.Collections.Generic;
using System.Text;

namespace GeneradorLiberaciones.Core.Utils;

public static class GradeTextConverter
{
    public static string CalificacionALetra(decimal? calificacion)
    {
        if (calificacion == null)
            return "NO APLICA";

        var valor = (int)calificacion.Value;

        if (valor < 70)
            return "REPROBADO";

        if (valor == 100)
            return "CIEN";

        // 70-99
        return NumeroATexto(valor);
    }

    // Versión simple, puedes afinarla después
    private static string NumeroATexto(int numero)
    {
        string[] decenas =
        {
            "", "DIEZ", "VEINTE", "TREINTA", "CUARENTA",
            "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"
        };

        string[] unidades =
        {
            "", "UNO", "DOS", "TRES", "CUATRO",
            "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"
        };

        if (numero < 10)
            return unidades[numero];

        int d = numero / 10;
        int u = numero % 10;

        if (u == 0)
            return decenas[d];

        return $"{decenas[d]} Y {unidades[u]}";
    }
}