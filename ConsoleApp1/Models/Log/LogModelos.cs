using System.Collections.Generic;

namespace ConsoleApp1.Models.Log
{
    public class LogProcesoGrifo
    {
        public string Grifo { get; set; } = string.Empty;
        public string Archivo { get; set; } = string.Empty;
        public List<LogDetalle> Mensajes { get; set; } = new List<LogDetalle>();
    }

    public class LogDetalle
    {
        public string Estado { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
    }
}
