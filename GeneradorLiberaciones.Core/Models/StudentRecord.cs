using System;
using System.Collections.Generic;
using System.Text;

namespace GeneradorLiberaciones.Core.Models;

public class StudentRecord
{
    public string Nombre { get; set; } = string.Empty;
    public string NoControl { get; set; } = string.Empty;
    public string Carrera { get; set; } = string.Empty;
    public string Dependencia { get; set; } = string.Empty;
    public string Programa { get; set; } = string.Empty;
    public DateTime? Inicio { get; set; }
    public DateTime? Termino { get; set; }
    public decimal? Calificacion { get; set; }
    public string FolioLiberacion { get; set; } = string.Empty;
    public string FolioCalificacion { get; set; } = string.Empty;
    public string CartaPresentacion { get; set; } = string.Empty;   // CARTA DE PRESENTACION
    public string CartaAceptacion { get; set; } = string.Empty;     // CARTA DE ACEPTACION
    public string CartaSolicitud { get; set; } = string.Empty;      // CARTA SOLICITUD
    public string CartaCompromiso { get; set; } = string.Empty;     // CARTA COMPROMISO
    public string Reporte1 { get; set; } = string.Empty;            // REPORTE 1
    public string Reporte2 { get; set; } = string.Empty;            // REPORTE 2
    public string Reporte3 { get; set; } = string.Empty;            // REPORTE 3
    public string CartaTerminacion { get; set; } = string.Empty;    // CARTA DE TERMINACION
    public string TrabajoFinal { get; set; } = string.Empty;        // TRABAJO FINAL
}
