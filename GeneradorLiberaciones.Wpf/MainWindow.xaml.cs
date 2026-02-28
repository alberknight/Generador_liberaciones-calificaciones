using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using GeneradorLiberaciones.Core.Services;

namespace GeneradorLiberaciones.Wpf
{
    public partial class MainWindow : Window
    {
        private FolioState folioState = new();
        private string outputRoot = string.Empty;
        private string foliosConfigPath = string.Empty;

        private readonly List<string> _logPaths = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var baseDir = AppContext.BaseDirectory;

            // La carpeta de salida ahora apunta a Documentos\GeneradorLiberaciones del usuario
            var documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            outputRoot = Path.Combine(documentsPath, "Resultados_Lib&Cali");
            
            // El archivo de configuración de folios se guarda en la raíz de la aplicación
            foliosConfigPath = Path.Combine(baseDir, "Folios.json");

            // Cargar FOLIOS + rutas desde JSON
            folioState = FolioStateStorage.Load(foliosConfigPath);

            int defaultLibStart = folioState.LastFolioLiberacion > 0
                ? folioState.LastFolioLiberacion + 1
                : 1;

            int defaultCalStart = folioState.LastFolioCalificacion > 0
                ? folioState.LastFolioCalificacion + 1
                : 1;

            txtFolioLib.Text = defaultLibStart.ToString();
            txtFolioCal.Text = defaultCalStart.ToString();

            // Excel por defecto
            if (!string.IsNullOrWhiteSpace(folioState.LastExcelPath) &&
                File.Exists(folioState.LastExcelPath))
            {
                txtExcelPath.Text = folioState.LastExcelPath;
            }

            // Plantillas por defecto
            var defaultTemplateLib = Path.Combine(AppContext.BaseDirectory, "Templates", "Liberacion.docx");
            var defaultTemplateCal = Path.Combine(AppContext.BaseDirectory, "Templates", "Calificacion.docx");

            txtTemplateLib.Text = !string.IsNullOrWhiteSpace(folioState.TemplateLiberacionPath)
                ? folioState.TemplateLiberacionPath
                : defaultTemplateLib;

            txtTemplateCal.Text = !string.IsNullOrWhiteSpace(folioState.TemplateCalificacionPath)
                ? folioState.TemplateCalificacionPath
                : defaultTemplateCal;

            // >>> AQUÍ se rellenan "Última ejecución" y "Últimos folios"
            UpdateLastRunInfo();

            AppendLog($"OutputRoot: {outputRoot}");
            AppendLog($"Último folio LIBERACIÓN: {folioState.LastFolioLiberacion}");
            AppendLog($"Último folio CALIFICACIÓN: {folioState.LastFolioCalificacion}");
        }



        private void UpdateLastRunInfo()
        {
            if (folioState.LastRunAt == default || folioState.LastRunAt.Year < 2000)
            {
                txtUltimaEjecucion.Text = "Sin registros previos";
            }
            else
            {
                txtUltimaEjecucion.Text = $"{folioState.LastRunAt:dd/MM/yyyy HH:mm:ss}";
            }

            txtUltimoFolioLib.Text = folioState.LastFolioLiberacion > 0
                ? folioState.LastFolioLiberacion.ToString()
                : "-";

            txtUltimoFolioCal.Text = folioState.LastFolioCalificacion > 0
                ? folioState.LastFolioCalificacion.ToString()
                : "-";
        }


        private void btnSeleccionarExcel_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*"
            };

            if (ofd.ShowDialog() == true)
            {
                txtExcelPath.Text = ofd.FileName;

                folioState.LastExcelPath = ofd.FileName;
                FolioStateStorage.Save(foliosConfigPath, folioState);
            }
        }

        private void btnSeleccionarTemplateLib_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Documentos Word (*.docx)|*.docx|Todos los archivos (*.*)|*.*"
            };

            if (ofd.ShowDialog() == true)
            {
                txtTemplateLib.Text = ofd.FileName;
                folioState.TemplateLiberacionPath = ofd.FileName;
                FolioStateStorage.Save(foliosConfigPath, folioState);
            }
        }

        private void btnSeleccionarTemplateCal_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Documentos Word (*.docx)|*.docx|Todos los archivos (*.*)|*.*"
            };

            if (ofd.ShowDialog() == true)
            {
                txtTemplateCal.Text = ofd.FileName;
                folioState.TemplateCalificacionPath = ofd.FileName;
                FolioStateStorage.Save(foliosConfigPath, folioState);
            }
        }


        private void btnAgregarLogs_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*",
                Multiselect = true
            };

            if (ofd.ShowDialog() == true)
            {
                foreach (var path in ofd.FileNames)
                {
                    if (!_logPaths.Contains(path))
                    {
                        _logPaths.Add(path);
                        lstLogs.Items.Add(path);
                    }
                }

                AppendLog($"Se agregaron {ofd.FileNames.Length} archivo(s) de log.");
            }
        }

        private void btnQuitarLog_Click(object sender, RoutedEventArgs e)
        {
            var selected = lstLogs.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selected))
                return;

            _logPaths.Remove(selected);
            lstLogs.Items.Remove(selected);
        }

        private async void btnGenerar_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnGenerar.IsEnabled = false;
                txtLog.Clear();

                var rutaExcel = txtExcelPath.Text.Trim();
                if (string.IsNullOrWhiteSpace(rutaExcel) || !File.Exists(rutaExcel))
                {
                    MessageBox.Show("Selecciona un archivo Excel válido.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // --- Plantillas ---
                var templateLiberacionPath = txtTemplateLib.Text.Trim();
                var templateCalificacionPath = txtTemplateCal.Text.Trim();

                if (string.IsNullOrWhiteSpace(templateLiberacionPath))
                    templateLiberacionPath = Path.Combine(AppContext.BaseDirectory, "Templates", "Liberacion.docx");

                if (string.IsNullOrWhiteSpace(templateCalificacionPath))
                    templateCalificacionPath = Path.Combine(AppContext.BaseDirectory, "Templates", "Calificacion.docx");

                if (!File.Exists(templateLiberacionPath))
                {
                    MessageBox.Show("No se encontró la plantilla de LIBERACIÓN.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!File.Exists(templateCalificacionPath))
                {
                    MessageBox.Show("No se encontró la plantilla de CALIFICACIÓN.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // === MODOS DE GENERACIÓN ===
                bool excluirYaGenerados = rbGenerarSoloNuevos.IsChecked == true;
                bool soloListado = rbGenerarSoloListado.IsChecked == true;

                // Logs:
                // - excluirYaGenerados -> NO_CONTROL previos (para excluir)
                // - soloListado        -> NO_CONTROL objetivo (solo éstos)
                var noControlPrevios = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var noControlObjetivo = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                bool modoUsaLogs = excluirYaGenerados || soloListado;

                if (modoUsaLogs)
                {
                    if (_logPaths.Count == 0)
                    {
                        MessageBox.Show(
                            "Has elegido un modo que requiere logs (.txt) pero no has cargado ninguno.",
                            "Atención",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                    else
                    {
                        foreach (var ruta in _logPaths)
                        {
                            var set = CargarNoControlDesdeLog(ruta);

                            if (excluirYaGenerados)
                                noControlPrevios.UnionWith(set);

                            if (soloListado)
                                noControlObjetivo.UnionWith(set);
                        }

                        if (excluirYaGenerados)
                            AppendLog($"Se cargaron {noControlPrevios.Count} NO_CONTROL para excluir desde {_logPaths.Count} archivo(s).");

                        if (soloListado)
                            AppendLog($"Se cargaron {noControlObjetivo.Count} NO_CONTROL objetivo (solo listado) desde {_logPaths.Count} archivo(s).");
                    }
                }

                if (!int.TryParse(txtFolioLib.Text.Trim(), out var folioInicialLib))
                {
                    MessageBox.Show("Folio inicial de liberaciones no es un número válido.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!int.TryParse(txtFolioCal.Text.Trim(), out var folioInicialCal))
                {
                    MessageBox.Show("Folio inicial de calificaciones no es un número válido.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var options = new GenerationOptions
                {
                    OutputRoot = outputRoot,
                    TemplateLiberacionPath = templateLiberacionPath,
                    TemplateCalificacionPath = templateCalificacionPath,
                    GenerarLiberaciones = true,
                    GenerarCalificaciones = true,
                    FolioInicialLiberacion = folioInicialLib,
                    FolioInicialCalificacion = folioInicialCal
                };

                string modoDesc = "Todos";
                if (excluirYaGenerados) modoDesc = "Solo nuevos (excluir NO_CONTROL ya generados)";
                if (soloListado) modoDesc = "Solo listado (generar solo NO_CONTROL del txt)";

                AppendLog("Iniciando generación...");
                AppendLog($"Excel: {rutaExcel}");
                AppendLog($"Modo: {modoDesc}");
                AppendLog($"Folio inicial LIB: {folioInicialLib}, FOL CAL: {folioInicialCal}");

                var excelReader = new ExcelReader();
                var templateEngine = new TemplateEngine();
                var service = new GenerationService(excelReader, templateEngine);

                // IMPORTANTE: debes adaptar la firma de GenerarTodo para aceptar estos parámetros nuevos
                var result = await System.Threading.Tasks.Task.Run(() =>
                    service.GenerarTodo(
                        rutaExcel,
                        options,
                        soloListado,
                        noControlObjetivo,
                        noControlPrevios,
                        excluirYaGenerados
                        ));

                AppendLog($"Documentos generados en esta ejecución: {result.DocumentosGenerados}");
                MessageBox.Show($"Documentos generados: {result.DocumentosGenerados}", "Proceso completado",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                // Si el modo es "solo listado", mostrar cuáles NO_CONTROL del txt no existen en el Excel
                if (soloListado && noControlObjetivo.Count > 0)
                {
                    var notFound = new List<string>();

                    foreach (var nc in noControlObjetivo)
                    {
                        // result.NoControlEncontradosEnExcel es un HashSet<string> llenado en GenerationService
                        if (!result.NoControlEncontradosEnExcel.Contains(nc))
                            notFound.Add(nc);
                    }

                    if (notFound.Count > 0)
                    {
                        AppendLog("Los siguientes NO_CONTROL del listado NO existen en este Excel:");
                        foreach (var nc in notFound)
                            AppendLog($"  - {nc}");
                    }
                    else
                    {
                        AppendLog("Todos los NO_CONTROL del listado existen en este Excel.");
                    }
                }

                if (result.LastFolioLiberacion.HasValue &&
                    result.LastFolioLiberacion.Value > folioState.LastFolioLiberacion)
                {
                    folioState.LastFolioLiberacion = result.LastFolioLiberacion.Value;
                }

                if (result.LastFolioCalificacion.HasValue &&
                    result.LastFolioCalificacion.Value > folioState.LastFolioCalificacion)
                {
                    folioState.LastFolioCalificacion = result.LastFolioCalificacion.Value;
                }

                // rutas y fecha de esta ejecución
                folioState.LastExcelPath = rutaExcel;
                folioState.TemplateLiberacionPath = templateLiberacionPath;
                folioState.TemplateCalificacionPath = templateCalificacionPath;
                folioState.LastRunAt = DateTime.Now;

                // guardar TODO al JSON
                FolioStateStorage.Save(foliosConfigPath, folioState);

                // refrescar UI
                UpdateLastRunInfo();

                AppendLog($"Último folio LIBERACIÓN guardado: {folioState.LastFolioLiberacion}");
                AppendLog($"Último folio CALIFICACIÓN guardado: {folioState.LastFolioCalificacion}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error durante la generación. Revisa el log.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                AppendLog("ERROR:");
                AppendLog(ex.ToString());
            }
            finally
            {
                btnGenerar.IsEnabled = true;
            }
        }


        private void AppendLog(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            txtLog.ScrollToEnd();
        }

        private static HashSet<string> CargarNoControlDesdeLog(string ruta)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                foreach (var line in File.ReadAllLines(ruta))
                {
                    var trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed))
                        continue;

                    if (trimmed.Equals("NO_CONTROL", StringComparison.OrdinalIgnoreCase))
                        continue;

                    set.Add(trimmed);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudo leer el archivo de log:\n{ruta}\n\n{ex.Message}",
                    "Advertencia", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            return set;
        }
    }
}
