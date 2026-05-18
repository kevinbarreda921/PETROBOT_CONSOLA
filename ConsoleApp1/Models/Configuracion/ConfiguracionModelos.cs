using System.Collections.Generic;

namespace ConsoleApp1.Models.Configuracion
{
    public class ConfigRoot
    {
        public Dictionary<string, GrifoConfig> Grifos { get; set; } = new();
    }

    public class GrifoConfig
    {
        public LecturaConfig? Lectura { get; set; }
        public EscrituraConfig? Escritura { get; set; }
        public Dictionary<string, string>? FilasClientesCreditos { get; set; }
    }

    public class LecturaConfig
    {
        public int ColumnaFecha { get; set; } = 14;
        public int ColumnaTotales { get; set; } = 15;
        public int ColumnaCreditoNombre { get; set; } = 0;
        public int ColumnaCreditoMonto { get; set; } = 6;
        public int ColumnaVariaCombusNombre { get; set; } = 16;
        public int ColumnaVariaCombusMonto { get; set; } = 18;
        public int ColumnaTablaHermes { get; set; } = 14;
        public Dictionary<string, string> MapeoFilas { get; set; } = new();
        public List<int> FilasVariaciones { get; set; } = new();
    }

    public class EscrituraConfig
    {
        public Dictionary<string, string> Columnas { get; set; } = new();
    }
}
