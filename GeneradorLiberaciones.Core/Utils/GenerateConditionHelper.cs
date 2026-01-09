using GeneradorLiberaciones.Core.Models;

namespace GeneradorLiberaciones.Core.Utils
{
    public static class GenerateConditionHelper
    {
        public static bool DebeGenerarDocumentos(StudentRecord record)
        {
            // Campos string obligatorios (desplegado para rendimiento)
            if (string.IsNullOrWhiteSpace(record.Nombre) ||
                string.IsNullOrWhiteSpace(record.NoControl) ||
                string.IsNullOrWhiteSpace(record.Dependencia) ||
                string.IsNullOrWhiteSpace(record.Programa) ||
                string.IsNullOrWhiteSpace(record.CartaPresentacion) ||
                string.IsNullOrWhiteSpace(record.CartaAceptacion) ||
                string.IsNullOrWhiteSpace(record.CartaSolicitud) ||
                string.IsNullOrWhiteSpace(record.CartaCompromiso) ||
                string.IsNullOrWhiteSpace(record.Reporte1) ||
                string.IsNullOrWhiteSpace(record.Reporte2) ||
                string.IsNullOrWhiteSpace(record.Reporte3) ||
                string.IsNullOrWhiteSpace(record.CartaTerminacion) ||
                string.IsNullOrWhiteSpace(record.TrabajoFinal))
            {
                return false;
            }

            // Fechas obligatorias
            if (record.Inicio == null || record.Termino == null)
                return false;

            // Calificación obligatoria
            if (record.Calificacion == null)
                return false;

            return true;
        }
    }
}
