using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace ConsoleApp1
{
    public class ExcelDataWrite
    {
        public class HojaGrifoMapeada
        {
            public string? Grifo { get; set; }
            public string? Hoja { get; set; }
            public Dictionary<string, int> MapaFechasFilas { get; set; } = new();
        }

        static ExcelDataWrite()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public List<HojaGrifoMapeada> MapearEstructuraMaestro(List<string> nombresGrifos)
        {
            string rutaMaestro = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
            string rutaExcel = Path.Combine(rutaMaestro, "ArchivosExcel", "Registro_ventas", "REGISTRO VENTAS -  2026- 01.xlsx");

            var resultados = new List<HojaGrifoMapeada>();

            using var stream = File.Open(rutaExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                string nombreHojaMin = reader.Name.ToLower();
                string? grifoMatch = nombresGrifos.FirstOrDefault(g => nombreHojaMin.Contains(g.ToLower()));

                if (grifoMatch != null)
                {
                    var mapeoHoja = new HojaGrifoMapeada { Grifo = grifoMatch, Hoja = reader.Name };
                    int filaActual = 0;

                    while (reader.Read())
                    {
                        filaActual++;

                        var valorRaw = reader.GetValue(1)?.ToString();

                        if (string.IsNullOrEmpty(valorRaw)) continue;

                        string valorCelda = valorRaw.Replace(" 00:00:00", "").Trim();

                        if (valorCelda.StartsWith("TOTAL", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (!mapeoHoja.MapaFechasFilas.ContainsKey(valorCelda))
                        {
                            mapeoHoja.MapaFechasFilas.Add(valorCelda, filaActual);
                        }
                    }
                    resultados.Add(mapeoHoja);
                }
            } while (reader.NextResult()); 

            return resultados;
        }


    }

}