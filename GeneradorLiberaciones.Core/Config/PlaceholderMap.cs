using System;
using System.Collections.Generic;
using System.Text;

using GeneradorLiberaciones.Core.Models;

namespace GeneradorLiberaciones.Core.Config;

public class PlaceholderMap
{
    public string Placeholder { get; }
    public Func<StudentRecord, string> Resolver { get; }

    public PlaceholderMap(string placeholder, Func<StudentRecord, string> resolver)
    {
        Placeholder = placeholder;
        Resolver = resolver;
    }
}
