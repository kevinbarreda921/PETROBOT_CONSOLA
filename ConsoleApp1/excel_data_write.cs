using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;
using ClosedXML.Excel;
using System.Text.Json;
using System.Reflection;

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



        public void EscribirFila(IXLWorksheet hoja, VentaDTO venta, int filaDestino, Dictionary<string, string> columnas)
        {
            var cachePropiedades = typeof(VentaDTO)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .ToDictionary(p => p.Name, p => p, StringComparer.Ordinal);

            foreach (var kvp in columnas)
            {
                string nombrePropiedad = kvp.Key;
                string columnaLetra = kvp.Value;

                if (!string.IsNullOrWhiteSpace(columnaLetra) && cachePropiedades.TryGetValue(nombrePropiedad, out PropertyInfo? propiedad))
                {
                    var valor = propiedad.GetValue(venta);
                    
                    // Lógica especial para Total_venta_acumulada solicitada
                    if (nombrePropiedad == "Total_venta_acumulada" && valor != null)
                    {
                        decimal totalVenta = 0;
                        if (valor is decimal decVal) totalVenta = decVal;
                        else decimal.TryParse(valor.ToString(), out totalVenta);

                        if (totalVenta != 0.0m)
                        {
                            string colGlp = columnas.TryGetValue("Venta_GPL", out string? cg) && !string.IsNullOrWhiteSpace(cg) ? $"{cg}{filaDestino}" : "0";
                            string colGnv = columnas.TryGetValue("Venta_GNV", out string? cn) && !string.IsNullOrWhiteSpace(cn) ? $"{cn}{filaDestino}" : "0";
                            
                            string strTotal = totalVenta.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            hoja.Cell($"{columnaLetra}{filaDestino}").FormulaA1 = $"{strTotal}-{colGlp}-{colGnv}";
                        }
                        continue;
                    }

                    if (valor != null)
                    {
                        // Convertir a string u otro tipo base si es necesario, ClosedXML maneja bien tipos primitivos.
                        if (valor is decimal decValor)
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = decValor;
                        else if (valor is double dblValor)
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = dblValor;
                        else if (valor is int intValor)
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = intValor;
                        else
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = valor.ToString();
                    }
                }
            }
        }

        public void EscribirClientesCredito(IXLWorksheet hoja, VentaDTO venta, int filaDestino, Dictionary<string, string> filasClientesCreditos, string grifoObjetivo)
        {
            if (venta.ListClienteCredito == null || venta.ListClienteCredito.Count == 0) return;

            // Invertir diccionario para buscar por Nombre de Cliente (Key = Nombre, Value = Columna Letra)
            var clienteAColumna = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in filasClientesCreditos)
            {
                if (!string.IsNullOrWhiteSpace(kvp.Value))
                {
                    clienteAColumna[kvp.Value.Trim()] = kvp.Key.Trim();
                }
            }

            foreach (var cliente in venta.ListClienteCredito)
            {
                string nombreLimpio = cliente.Cliente?.Trim() ?? "";
                if (string.IsNullOrEmpty(nombreLimpio)) continue;

                if (clienteAColumna.TryGetValue(nombreLimpio, out string? columnaLetra) && !string.IsNullOrWhiteSpace(columnaLetra))
                {
                    var valor = cliente.Monto;
                    if (valor != null)
                    {
                        if (valor is decimal decValor)
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = decValor;
                        else if (valor is double dblValor)
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = dblValor;
                        else if (valor is int intValor)
                            hoja.Cell($"{columnaLetra}{filaDestino}").Value = intValor;
                        else
                        {
                            if (decimal.TryParse(valor.ToString(), out decimal parsedDec))
                                hoja.Cell($"{columnaLetra}{filaDestino}").Value = parsedDec;
                            else
                                hoja.Cell($"{columnaLetra}{filaDestino}").Value = valor.ToString();
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"[!] El cliente '{nombreLimpio}' no existe en la configuración JSON del grifo {grifoObjetivo}");
                }
            }
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

                        string valorCelda = valorRaw.Replace("12:00:00 a. m.", "").Trim();

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