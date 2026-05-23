using ConsoleApp1.Models;
using ConsoleApp1.Models.Configuracion;
using ExcelDataReader;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Services
{
    public class LectorExcelService
    {
        static LectorExcelService()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public ConcurrentBag<ArchivoGrifo> LeerPartesDiarios()
        {
            string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

            if (!Directory.Exists(carpeta)) return new ConcurrentBag<ArchivoGrifo>();

            var archivos = Directory.GetFiles(carpeta, "*.xlsx");
            var listaGrifosProcesar = new ConcurrentBag<ArchivoGrifo>();

            Dictionary<string, PropertyInfo> cachePropiedades = typeof(VentaDTO)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanWrite)
                .ToDictionary(p => p.Name, p => p, StringComparer.Ordinal);

            var listadoClavesGrifos = ConfiguracionService.ConfigGlobal.Grifos.Keys.ToList();

            Parallel.ForEach(archivos, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, (ruta) =>
            {
                string nombreGrifoDetectadoStr = "DESCONOCIDO";
                try
                {
                    string nombreArchivoCompleto = Path.GetFileNameWithoutExtension(ruta);
                    string nombreArchivoMin = nombreArchivoCompleto.ToLower();

                    string? nombreGrifoDetectado = listadoClavesGrifos.FirstOrDefault(clave => nombreArchivoMin.Contains(clave.ToLower()));

                    if (nombreGrifoDetectado == null)
                    {
                        LoggerService.Error("DESCONOCIDO", Path.GetFileName(ruta), $"NO REGISTRADO: El archivo a procesar no se encuentra registrado en el sistema.");
                        return;
                    }
                    
                    nombreGrifoDetectadoStr = nombreGrifoDetectado;

                    var configGrifoRoot = ConfiguracionService.ConfigGlobal.Grifos[nombreGrifoDetectado];
                    var configGrifo = configGrifoRoot.Lectura;

                    if (configGrifo == null) return;

                    var filasDeseadas = new HashSet<int>(configGrifo.MapeoFilas.Keys.Select(int.Parse));

                    ArchivoGrifo nuevoGrifo = new ArchivoGrifo(nombreGrifoDetectado, Path.GetFileName(ruta));

                    using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan);
                    using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    do
                    {
                        var registro = new VentaDTO { Hoja = reader.Name };
                        decimal descuentoLiquidos_Total = 0;
                        int filaActual = 1;
                        bool leyendoClientes = false;
                        bool leyendoVariaciones = false;
                        int flagclientecredito = 0;
                        var clientesAgrupados = new Dictionary<string, decimal>();

                        bool hermesTablaEncontrada = false;
                        bool hermesLeyendoTablaFlotante = false;
                        string hermesPalabraClaveCabecera = "IMPORTE S/.";
                        var listaHermes = new List<(string Banco, string Tipo, decimal Monto)>();

                        while (reader.Read())
                        {
                            if (filaActual == 3)
                            {
                                int colLetraTotales = configGrifo.ColumnaFecha;
                                var fecha_hoja = reader.GetValue(colLetraTotales);
                                if (fecha_hoja != null)
                                {
                                    if (fecha_hoja is DateTime dt) registro.Dia = dt.ToString("d/MM/yyyy");
                                    else if (DateTime.TryParse(fecha_hoja.ToString(), out DateTime parsedDate)) registro.Dia = parsedDate.ToString("d/MM/yyyy");
                                    else
                                    {
                                        string rawFecha = fecha_hoja.ToString() ?? "";
                                        int indexSpace = rawFecha.IndexOf(" 00:00");
                                        registro.Dia = indexSpace != -1 ? rawFecha.Substring(0, indexSpace) : rawFecha.Trim();
                                    }
                                }
                            }

                            if (filaActual > 129) break;

                            if (filasDeseadas.Contains(filaActual))
                            {
                                int colLetraTotales = configGrifo.ColumnaTotales;
                                var valor = reader.GetValue(colLetraTotales);
                                decimal numValor = 0;
                                if (valor != null)
                                {
                                    string cleanStr = (valor.ToString() ?? "").Replace(",", "").Replace("-", "");
                                    decimal.TryParse(cleanStr, out numValor);
                                }

                                if (configGrifo.MapeoFilas.TryGetValue(filaActual.ToString(), out string? nombrePropiedad))
                                {
                                    if (cachePropiedades.TryGetValue(nombrePropiedad, out PropertyInfo? propiedad))
                                    {
                                        propiedad.SetValue(registro, numValor);
                                    }
                                }
                            }

                            if (!leyendoClientes && flagclientecredito == 0 && filaActual >= 10 && filaActual <= 30)
                            {
                                int colLetraCreditoNombre = configGrifo.ColumnaCreditoNombre;
                                var valorNombre = reader.GetValue(colLetraCreditoNombre);
                                if (valorNombre != null && valorNombre.ToString()?.Trim().ToUpper().StartsWith("CLIENTE") == true)
                                {
                                    leyendoClientes = true;
                                }
                            }
                            else if (leyendoClientes)
                            {
                                int colLetraCreditoNombre = configGrifo.ColumnaCreditoNombre;
                                int colLetraCreditoMonto = configGrifo.ColumnaCreditoMonto;
                                var valorNombre = reader.GetValue(colLetraCreditoNombre);
                                var valorMonto = reader.GetValue(colLetraCreditoMonto);

                                if (valorNombre != null && !string.IsNullOrWhiteSpace(valorNombre.ToString()))
                                {
                                    string nombreLimpio = (valorNombre.ToString() ?? "").Trim();
                                    decimal.TryParse(valorMonto?.ToString(), out decimal montoActual);

                                    if (nombreGrifoDetectadoStr == "ACAPULCO") { 
                                       nombreLimpio = nombreLimpio.Length > 20 ? nombreLimpio.Substring(0, 20) : nombreLimpio;
                                    }
                                    if (clientesAgrupados.TryGetValue(nombreLimpio, out decimal montoExistente))
                                        clientesAgrupados[nombreLimpio] = montoExistente + montoActual;
                                    else
                                        clientesAgrupados.Add(nombreLimpio, montoActual);

                                    flagclientecredito = 0;
                                }
                                else
                                {
                                    if (flagclientecredito == 1) leyendoClientes = false;
                                    flagclientecredito = 1;
                                }
                            }

                            if (filaActual >= 17 && filaActual <= 50)
                            {
                                int colLetraColumnaVariaCombusNombre = configGrifo.ColumnaVariaCombusNombre;
                                var varia_combus_nombre = reader.GetValue(colLetraColumnaVariaCombusNombre);
                                string? nombreTrimmed = varia_combus_nombre?.ToString()?.Trim();

                                if (leyendoVariaciones)
                                {
                                    if (string.Equals(nombreTrimmed, "TOTAL", StringComparison.OrdinalIgnoreCase))
                                    {
                                        leyendoVariaciones = false;
                                    }
                                    else
                                    {
                                        int colLetraColumnaVariaCombusMonto = configGrifo.ColumnaVariaCombusMonto;
                                        var varia_combus_monto = reader.GetValue(colLetraColumnaVariaCombusMonto)?.ToString()?.Replace("-", "");
                                           
                                        decimal.TryParse(varia_combus_monto, out decimal montoActualVariacion);

                                        if (string.Equals(nombreTrimmed, "GLP", StringComparison.OrdinalIgnoreCase))
                                            registro.DescuentoGLP = montoActualVariacion;
                                        else
                                            descuentoLiquidos_Total += montoActualVariacion;
                                    }
                                }
                                else
                                {
                                    if (string.Equals(nombreTrimmed, "COMBUSTIBLE", StringComparison.OrdinalIgnoreCase))
                                    {
                                        leyendoVariaciones = true;
                                    }
                                }
                            }

                            if (filaActual >= 40 && filaActual <= 130)
                            {
                                int colLetraColumnaTablaHermes = configGrifo.ColumnaTablaHermes;
                                var celdaIdentificadora = reader.GetValue(colLetraColumnaTablaHermes);
                                string textoCelda = celdaIdentificadora?.ToString()?.Trim() ?? "";

                                if (!hermesTablaEncontrada && textoCelda.Contains(hermesPalabraClaveCabecera, StringComparison.OrdinalIgnoreCase))
                                {
                                    hermesTablaEncontrada = true;
                                    hermesLeyendoTablaFlotante = true;
                                    filaActual++;
                                    continue;
                                }

                                if (hermesLeyendoTablaFlotante)
                                {
                                    if (!string.IsNullOrWhiteSpace(textoCelda))
                                    {
                                        decimal.TryParse(textoCelda, out decimal montoActualHermes);

                                        var celdahermesbanco = reader.GetValue(colLetraColumnaTablaHermes - 5);
                                        string textohermesbanco = celdahermesbanco?.ToString()?.Trim() ?? "";

                                        var celdahermestipo = reader.GetValue(colLetraColumnaTablaHermes + 2);
                                        string textohermestipo = celdahermestipo?.ToString()?.Trim() ?? "";

                                        listaHermes.Add((textohermesbanco, textohermestipo, montoActualHermes));
                                    }
                                    else
                                    {
                                        hermesLeyendoTablaFlotante = false;
                                        break; // Ya no recorrer mas filas
                                    }
                                }
                            }

                            filaActual++;
                        }

                        foreach (var entrada in clientesAgrupados)
                        {
                            registro.AgregarClienteCredito(entrada.Key, entrada.Value);
                        }
                        registro.DescuentoLiquidos = descuentoLiquidos_Total;

                        var listaFiltrada = listaHermes
                            .Where(x => string.Equals(x.Banco, "SCOTIABANK", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(x.Banco, "MIBANCO", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        var liquidosScotia = listaFiltrada
                            .Where(x => x.Tipo.Contains("liquido", StringComparison.OrdinalIgnoreCase) || 
                                        x.Tipo.Contains("líquido", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        var glpScotia = listaFiltrada
                            .Where(x => x.Tipo.Contains("GLP", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        var gnvScotia = listaFiltrada
                            .Where(x => x.Tipo.Contains("GNV", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        // Validaciones de límites
                        bool cumpleReglas = true;
                        string mensajeError = "";

                        if (liquidosScotia.Count > 1)
                        {
                            cumpleReglas = false;
                            mensajeError += $"Se encontraron {liquidosScotia.Count} registros de Liquido para SCOTIABANK (maximo permitido: 1). ";
                        }

                        if (glpScotia.Count > 1)
                        {
                            cumpleReglas = false;
                            mensajeError += $"Se encontraron {glpScotia.Count} registros de GLP para SCOTIABANK (maximo permitido: 1). ";
                        }

                        if (gnvScotia.Count > 2)
                        {
                            cumpleReglas = false;
                            mensajeError += $"Se encontraron {gnvScotia.Count} registros de GNV para SCOTIABANK (maximo permitido: 2). ";
                        }

                        if (cumpleReglas)
                        {
                            // Llenado de variables si cumple las reglas
                            registro.Hermes_monto_liquido = liquidosScotia.Any() ? liquidosScotia.Sum(x => x.Monto) : 0m;
                            registro.Hermes_monto_GLP = glpScotia.Any() ? glpScotia.Sum(x => x.Monto) : 0m;

                            if (gnvScotia.Count == 1)
                            {
                                registro.Hermes_monto_GNV1 = gnvScotia[0].Monto;
                                registro.Hermes_monto_GNV2 = 0m;
                            }
                            else if (gnvScotia.Count >= 2)
                            {
                                var ordenadosGnv = gnvScotia.OrderByDescending(x => x.Monto).ToList();
                                registro.Hermes_monto_GNV1 = ordenadosGnv.First().Monto;
                                registro.Hermes_monto_GNV2 = ordenadosGnv.Last().Monto;
                            }
                            else
                            {
                                registro.Hermes_monto_GNV1 = 0m;
                                registro.Hermes_monto_GNV2 = 0m;
                            }
                        }
                        else
                        {
                            // Registrar error en reporte_proceso.json y poner variables en 0
                            LoggerService.Error(nombreGrifoDetectadoStr, Path.GetFileName(ruta), $"En el dia {registro.Dia} {mensajeError}");

                            registro.Hermes_monto_liquido = 0m;
                            registro.Hermes_monto_GLP = 0m;
                            registro.Hermes_monto_GNV1 = 0m;
                            registro.Hermes_monto_GNV2 = 0m;
                        }

                        nuevoGrifo.AgregarVenta(registro);

                    } while (reader.NextResult());

                    listaGrifosProcesar.Add(nuevoGrifo);
                }
                catch (Exception ex)
                {
                    LoggerService.Error(nombreGrifoDetectadoStr, Path.GetFileName(ruta), $"Error leyendo el archivo: {ex.Message}");
                }
            });

            return listaGrifosProcesar;
        }
    }
}
