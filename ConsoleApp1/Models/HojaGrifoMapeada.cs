using System.Collections.Generic;

namespace ConsoleApp1.Models
{
    public class HojaGrifoMapeada
    {
        public string? Grifo { get; set; }
        public string? Hoja { get; set; }
        public Dictionary<string, int> MapaFechasFilas { get; set; } = new();
    }
}
