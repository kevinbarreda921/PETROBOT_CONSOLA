using OfficeOpenXml;
using ConsoleApp1.Models;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ConsoleApp1.Services
{
    public class EscritorExcelService
    {
        private static readonly Dictionary<string, PropertyInfo> _cachePropiedades = typeof(VentaDTO)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .ToDictionary(p => p.Name, p => p, StringComparer.Ordinal);

        public void EscribirFila(ExcelWorksheet hoja, VentaDTO venta, int filaDestino, Dictionary<string, string> columnas)
        {

            foreach (var kvp in columnas)
            {
                string nombrePropiedad = kvp.Key;
                string columnaLetra = kvp.Value;

                if (!string.IsNullOrWhiteSpace(columnaLetra) && _cachePropiedades.TryGetValue(nombrePropiedad, out PropertyInfo? propiedad))
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
                            hoja.Cells[$"{columnaLetra}{filaDestino}"].Formula = $"{strTotal}-{colGlp}-{colGnv}";
                        }
                        continue;
                    }

                    if (valor != null)
                    {
                        bool esCero = false;
                        if (valor is decimal dec && dec == 0m) esCero = true;
                        else if (valor is double dbl && dbl == 0d) esCero = true;
                        else if (valor is int integer && integer == 0) esCero = true;
                        else if (decimal.TryParse(valor.ToString(), out decimal parsed) && parsed == 0m) esCero = true;

                        if (!esCero)
                        {
                            if (valor is decimal decValor)
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = decValor;
                            else if (valor is double dblValor)
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = dblValor;
                            else if (valor is int intValor)
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = intValor;
                            else
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = valor.ToString();
                        }
                    }
                }
            }
        }

        public void EscribirClientesCredito(ExcelWorksheet hoja, VentaDTO venta, int filaDestino, Dictionary<string, string> clienteAColumna, string grifoObjetivo, string archivo)
        {
            if (venta.ListClienteCredito == null || venta.ListClienteCredito.Count == 0) return;

            foreach (var cliente in venta.ListClienteCredito)
            {
                string nombreLimpio = cliente.Cliente?.Trim() ?? "";
                if (string.IsNullOrEmpty(nombreLimpio)) continue;

                if (clienteAColumna.TryGetValue(nombreLimpio, out string? columnaLetra) && !string.IsNullOrWhiteSpace(columnaLetra))
                {
                    var valor = cliente.Monto;
                    if (valor != null)
                    {
                        bool esCero = false;
                        if (valor is decimal dec && dec == 0m) esCero = true;
                        else if (valor is double dbl && dbl == 0d) esCero = true;
                        else if (valor is int integer && integer == 0) esCero = true;
                        else if (decimal.TryParse(valor.ToString(), out decimal parsed) && parsed == 0m) esCero = true;

                        if (!esCero)
                        {
                            if (valor is decimal decValor)
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = decValor;
                            else if (valor is double dblValor)
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = dblValor;
                            else if (valor is int intValor)
                                hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = intValor;
                            else
                            {
                                if (decimal.TryParse(valor.ToString(), out decimal parsedDec))
                                    hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = parsedDec;
                                else
                                    hoja.Cells[$"{columnaLetra}{filaDestino}"].Value = valor.ToString();
                            }
                        }
                    }
                }
                else
                {
                    LoggerService.Error(grifoObjetivo, archivo, $"El cliente '{nombreLimpio}' no existe en la configuración JSON");
                }
            }
        }

        public Dictionary<string, int> MapearFechasHoja(ExcelWorksheet hoja)
        {
            var mapaFechasFilas = new Dictionary<string, int>();

            if (hoja.Dimension == null) return mapaFechasFilas;

            int maxRow = hoja.Dimension.End.Row;

            for (int filaActual = 1; filaActual <= maxRow; filaActual++)
            {
                var valorRaw = hoja.Cells[filaActual, 2].Value?.ToString();

                if (string.IsNullOrEmpty(valorRaw)) continue;

                // string valorCelda = valorRaw.Replace("12:00:00 a. m.", "").Trim(); //trabajo
                string valorCelda = valorRaw.Replace("00:00:00", "").Trim(); //casas

                if (valorCelda.StartsWith("TOTAL", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!mapaFechasFilas.ContainsKey(valorCelda))
                {
                    mapaFechasFilas.Add(valorCelda, filaActual);
                }
            }

            return mapaFechasFilas;
        }
    }
}
