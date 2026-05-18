using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1
{
    public class ArchivoGrifo
    {
        public string? Grifo { get; set; }
        public string? Archivo { get; set; }
        // CORRECCIÓN CLAVE: Al agregar `= new();`, la lista se crea automáticamente 
        // en el segundo exacto en que el objeto nace en memoria. Nunca más será null.
        public List<VentaDTO> ListVenta { get; set; } = new();

        // Tu constructor con parámetros
        public ArchivoGrifo(string? grifo, string? archivo)
        {
            Grifo = grifo;
            Archivo = archivo;
        }

        // Tu constructor vacío
        public ArchivoGrifo()
        {
        }

        public void AgregarVenta(VentaDTO venta)
        {
            if (venta != null)
            {
                ListVenta.Add(venta); // Ahora funcionará siempre sin romperse
            }
        }
    }
}