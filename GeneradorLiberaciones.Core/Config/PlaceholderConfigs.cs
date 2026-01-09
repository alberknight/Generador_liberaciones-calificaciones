using System;
using System.Collections.Generic;
using System.Text;
using GeneradorLiberaciones.Core.Models;
using GeneradorLiberaciones.Core.Utils;

namespace GeneradorLiberaciones.Core.Config;

public static class PlaceholderConfigs
{
    public static IReadOnlyList<PlaceholderMap> Liberacion => new[]
    {
        new PlaceholderMap("NOMBRE_ALUMNO", r => r.Nombre),
        new PlaceholderMap("NO_DE_CONTROL", r => r.NoControl),
        new PlaceholderMap("CARRERA_ALUMNO", r => r.Carrera),
        new PlaceholderMap("DEPENDENCIA_ALUMNO", r => r.Dependencia),
        new PlaceholderMap("PROGRAMA_ALUMNO", r => r.Programa),
        new PlaceholderMap("NO_DE_FOLIO", r => r.FolioLiberacion),
        new PlaceholderMap("FECHA_INICIO", r => DateUtils.FormatearFechaLarga(r.Inicio) ?? string.Empty),
        new PlaceholderMap("FECHA_TERMINO", r => DateUtils.FormatearFechaLarga(r.Termino) ?? string.Empty),
    };

    public static IReadOnlyList<PlaceholderMap> Calificacion => new[]
    {
        new PlaceholderMap("NOMBRE_ALUMNO", r => r.Nombre),
        new PlaceholderMap("NO_DE_CONTROL", r => r.NoControl),
        new PlaceholderMap("CARRERA_ALUMNO", r => r.Carrera),
        new PlaceholderMap("CALIFICACION_NUMERO", r => r.Calificacion?.ToString("0") ?? string.Empty),
        new PlaceholderMap("CALIFICACION_LETRA", r => GradeTextConverter.CalificacionALetra(r.Calificacion)),
        new PlaceholderMap("NO_DE_FOLIO", r => r.FolioCalificacion),
    };
}