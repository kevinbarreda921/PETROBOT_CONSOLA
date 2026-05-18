using ConsoleApp1;
using ExcelDataReader;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text;
using System.Text.Json;

public class excel_data_reader_parallel
{
    public class ConfigRoot
    {
        public Dictionary<string, GrifoConfig> Grifos { get; set; } = new();
    }

    public class GrifoConfig
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

    private static ConfigRoot _configGlobal = new();

    static excel_data_reader_parallel()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        CargarConfiguracionJson();
    }

    private static void CargarConfiguracionJson()
    {
        string rutaConfig = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\", "config_grifos.json"));

        if (File.Exists(rutaConfig))
        {
            try
            {
                string jsonTexto = File.ReadAllText(rutaConfig);
                _configGlobal = JsonSerializer.Deserialize<ConfigRoot>(jsonTexto) ?? new ConfigRoot();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error crítico al leer el JSON de configuración: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"Advertencia: No se encontró el archivo de configuración en: {rutaConfig}");
        }
    }

    public static ConcurrentBag<ArchivoGrifo> f_excel_data_reader_parallel()
    {
        string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
        string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

        if (!Directory.Exists(carpeta)) return new ConcurrentBag<ArchivoGrifo>();

        var archivos = Directory.GetFiles(carpeta, "*.xlsx");
        var listaGrifosProcesar = new ConcurrentBag<ArchivoGrifo>();

        // Cacheamos las propiedades por reflexión para alto rendimiento
        Dictionary<string, PropertyInfo> cachePropiedades = typeof(VentaDTO)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite)
            .ToDictionary(p => p.Name, p => p, StringComparer.Ordinal);

        // Extraemos el listado de llaves configuradas en el JSON
        var listadoClavesGrifos = _configGlobal.Grifos.Keys.ToList();

        Parallel.ForEach(archivos, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, (ruta) =>
        {
            try
            {
                string nombreArchivoCompleto = Path.GetFileNameWithoutExtension(ruta);
                string nombreArchivoMin = nombreArchivoCompleto.ToLower();

                // Búsqueda difusa para detectar qué grifo pertenece al archivo analizado
                string? nombreGrifoDetectado = listadoClavesGrifos.FirstOrDefault(clave => nombreArchivoMin.Contains(clave.ToLower()));

                if (nombreGrifoDetectado == null)
                {
                    Console.WriteLine($"[❌ ERROR] NO REGISTRADO: El archivo '{Path.GetFileName(ruta)}' no coincide con ninguna palabra clave del JSON.");
                    return;
                }

                // Recuperamos la configuración correspondiente de la clave detectada
                var configGrifo = _configGlobal.Grifos[nombreGrifoDetectado];

                // Construimos los HashSets dinámicos
                var filasDeseadas = new HashSet<int>(configGrifo.MapeoFilas.Keys.Select(int.Parse));
                var filasVariaciones = new HashSet<int>(configGrifo.FilasVariaciones);

                ArchivoGrifo nuevoGrifo = new ArchivoGrifo(nombreGrifoDetectado, Path.GetFileName(ruta));

                using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan);
                using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                do
                {
                    var registro = new VentaDTO
                    {
                        Hoja = reader.Name
                    };

                    decimal descuentoLiquidos_Total = 0;
                    int filaActual = 1;
                    bool leyendoClientes = false;
                    int flagclientecredito = 0;
                    var clientesAgrupados = new Dictionary<string, decimal>();

                    bool hermesTablaEncontrada = false;
                    bool hermesLeyendoTablaFlotante = false;
                    string hermesPalabraClaveCabecera = "IMPORTE S/.";
                    int contadorHermes = 1;

                    while (reader.Read())
                    {
                        // Fila 3: Conversión Segura de Fechas
                        if (filaActual == 3)
                        {
                            int colLetraTotales = configGrifo.ColumnaFecha;
                            var fecha_hoja = reader.GetValue(colLetraTotales);
                            if (fecha_hoja != null)
                            {
                                if (fecha_hoja is DateTime dt) registro.Dia = dt.ToString("dd/MM/yyyy");
                                else if (DateTime.TryParse(fecha_hoja.ToString(), out DateTime parsedDate)) registro.Dia = parsedDate.ToString("dd/MM/yyyy");
                                else
                                {
                                    string rawFecha = fecha_hoja.ToString() ?? "";
                                    int indexSpace = rawFecha.IndexOf(" 00:00");
                                    registro.Dia = indexSpace != -1 ? rawFecha.Substring(0, indexSpace) : rawFecha.Trim();
                                }
                            }
                        }

                        if (filaActual > 129 || contadorHermes == 5) break;

                        // Mapeo Dinámico
                        if (filasDeseadas.Contains(filaActual))
                        {
                            // Se utiliza el valor numérico dinámico proveniente del JSON (ej: 15)
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

                        // Clientes crédito
                        if (filaActual == 15 || leyendoClientes)
                        {

                            int colLetraCreditoNombre = configGrifo.ColumnaCreditoNombre;
                            int colLetraCreditoMonto= configGrifo.ColumnaCreditoMonto;
                            var valorNombre = reader.GetValue(colLetraCreditoNombre);
                            var valorMonto = reader.GetValue(colLetraCreditoMonto);

                            if (valorNombre != null && !string.IsNullOrWhiteSpace(valorNombre.ToString()))
                            {
                                string nombreLimpio = (valorNombre.ToString() ?? "").Trim();
                                decimal.TryParse(valorMonto?.ToString(), out decimal montoActual);

                                if (clientesAgrupados.TryGetValue(nombreLimpio, out decimal montoExistente))
                                    clientesAgrupados[nombreLimpio] = montoExistente + montoActual;
                                else
                                    clientesAgrupados.Add(nombreLimpio, montoActual);

                                leyendoClientes = true;
                            }
                            else
                            {
                                if (flagclientecredito == 1) leyendoClientes = false;
                                flagclientecredito = 1;
                            }
                        }

                        // Variaciones
                        if (filasVariaciones.Contains(filaActual))
                        {

                            int colLetraColumnaVariaCombusNombre = configGrifo.ColumnaVariaCombusNombre;
                            int colLetraColumnaVariaCombusMonto  = configGrifo.ColumnaVariaCombusMonto;
                            var varia_combus_nombre = reader.GetValue(colLetraColumnaVariaCombusNombre);
                            var varia_combus_monto = reader.GetValue(colLetraColumnaVariaCombusMonto);
                            decimal.TryParse(varia_combus_monto?.ToString(), out decimal montoActualVariacion);

                            if (varia_combus_nombre?.ToString()?.Trim() == "GLP")
                                registro.DescuentoGLP = montoActualVariacion;
                            else
                                descuentoLiquidos_Total += montoActualVariacion;
                        }

                        // Tabla Flotante de Hermes
                        if (filaActual >= 60 && filaActual <= 130)
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

                                    switch (contadorHermes)
                                    {
                                        case 1: registro.Hermes_monto_liquido = montoActualHermes; break;
                                        case 2: registro.Hermes_monto_GLP = montoActualHermes; break;
                                        case 3: registro.Hermes_monto_GNV1 = montoActualHermes; break;
                                        case 4: registro.Hermes_monto_GNV2 = montoActualHermes; break;
                                    }
                                    contadorHermes++;
                                }
                                else
                                {
                                    hermesLeyendoTablaFlotante = false;
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

                    nuevoGrifo.AgregarVenta(registro);

                } while (reader.NextResult());

                listaGrifosProcesar.Add(nuevoGrifo);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error procesando el archivo {Path.GetFileName(ruta)}: {ex.Message}");
            }
        });
        return listaGrifosProcesar;
    }
}