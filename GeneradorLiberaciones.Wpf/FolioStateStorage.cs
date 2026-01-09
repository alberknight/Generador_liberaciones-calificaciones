using System;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace GeneradorLiberaciones.Wpf
{
    internal class FolioState
    {
        public int LastFolioLiberacion { get; set; }
        public int LastFolioCalificacion { get; set; }
        public DateTime LastRunAt { get; set; }

        public string? LastExcelPath { get; set; }
        public string? TemplateLiberacionPath { get; set; }
        public string? TemplateCalificacionPath { get; set; }
    }

    internal static class FolioStateStorage
    {
        public static FolioState Load(string path)
        {
            try
            {
                if (!File.Exists(path))
                    return new FolioState();

                var json = File.ReadAllText(path);
                var state = JsonSerializer.Deserialize<FolioState>(json);
                return state ?? new FolioState();
            }
            catch
            {
                // Si algo falla, devolvemos un estado vacío
                return new FolioState();
            }
        }

        public static void Save(string path, FolioState state)
        {
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var json = JsonSerializer.Serialize(
                    state,
                    new JsonSerializerOptions
                    {
                        WriteIndented = true
                    });

                File.WriteAllText(path, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"No se pudo guardar el archivo de folios:\n{ex.Message}",
                    "Advertencia",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
    }
}
