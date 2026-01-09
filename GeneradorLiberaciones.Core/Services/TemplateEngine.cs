using System;
using System.Collections.Generic;
using System.Text;
using GeneradorLiberaciones.Core.Config;
using GeneradorLiberaciones.Core.Models;
using Xceed.Words.NET;

namespace GeneradorLiberaciones.Core.Services;

public class TemplateEngine
{
    public void GenerarDocumento(
        string templatePath,
        string outputPath,
        StudentRecord record,
        IReadOnlyList<PlaceholderMap> maps)
    {
        using var doc = DocX.Load(templatePath);
            
        foreach (var map in maps)
        {
            var value = map.Resolver(record) ?? string.Empty;
            doc.ReplaceText(map.Placeholder, value);
        }

        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        doc.SaveAs(outputPath);
    }
}
