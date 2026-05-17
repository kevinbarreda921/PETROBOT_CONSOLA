using ConsoleApp1;
using ExcelDataReader;
using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;

public class excel_data_reader_parallel
{
    // 1. Registro del Encoding movido al constructor estático para que se ejecute solo una vez
    static excel_data_reader_parallel()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static void f_excel_data_reader_parallel()
    {
        string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
        string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

        if (!Directory.Exists(carpeta)) return;

        var archivos = Directory.GetFiles(carpeta, "*.xlsx");
        var listaResultados = new ConcurrentBag<VentaDTO>();
        var filasDeseadas = new HashSet<int> { 8, 9, 12, 17, 18, 19, 26, 28, 29, 31 };
        var FilasVariaciones = new HashSet<int> { 21, 22, 23, 24, 25 };
        var FilasHermes = new HashSet<int> { 125, 126, 127, 128 };

        Parallel.ForEach(archivos, (ruta) =>
        {
            try
            {
                using var stream = File.Open(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    var registro = new VentaDTO
                    {
                        Archivo = Path.GetFileName(ruta),
                        Hoja = reader.Name
                    };
                    decimal DescuentoLiquidos_Total = 0;

                    int filaActual = 1;
                    bool leyendoClientes = false;
                    int flagclientecredito = 0;
                    Dictionary<string, decimal> clientesAgrupados = new Dictionary<string, decimal>();

                    bool HermesTablaEncontra = false;
                    bool HermesLeyendoTablaFlotante = false;
                    string HermesPalabraClaveCabecera = "IMPORTE S/.";

                    int ContadorHermes = 1;

                    while (reader.Read())
                    {
                        // 2. Conversión Segura de Fechas
                        if (filaActual == 3)
                        {
                            var fecha_hoja = reader.GetValue(14); // Columna O
                            if (fecha_hoja != null)
                            {
                                if (fecha_hoja is DateTime dt)
                                {
                                    registro.Dia = dt.ToString("dd/MM/yyyy");
                                }
                                else if (DateTime.TryParse(fecha_hoja.ToString(), out DateTime parsedDate))
                                {
                                    registro.Dia = parsedDate.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    registro.Dia = fecha_hoja.ToString().Replace(" 00:00:00", "").Trim();
                                }
                            }
                        }

                        if (filaActual > 129 || ContadorHermes == 5) break;

                        // 3. Asignación con Manejo de Nulos
                        if (filasDeseadas.Contains(filaActual))
                        {
                            var valor = reader.GetValue(15); // Columna P

                            string valorStr = valor?.ToString()?.Replace(",", "") ?? "0";
                            string valorTarjeta = valor?.ToString()?.Replace("-", "")?.Replace(",", "") ?? "0";

                            switch (filaActual)
                            {
                                case 8: registro.Venta_GPL = valorStr; break;
                                case 9: registro.Venta_GNV = valorStr; break;
                                case 12: registro.Total_venta_acumulada = valorStr; break;
                                case 17: registro.Total_Tarjeta_de_Credito_Liquidos = valorTarjeta; break;
                                case 18: registro.Total_Tarjeta_de_Credito_GLP = valorTarjeta; break;
                                case 19: registro.Total_Tarjeta_de_Credito_GNV = valorTarjeta; break;
                                case 26: registro.ErrorMaquina = valor ?? 0; break;
                                case 28: registro.Recaudo_Cofide_GNV = valorStr; break;
                                case 29: registro.Gastos = valor ?? 0; break;
                                case 31: registro.Ventas_con_transferencia = valor ?? 0; break;
                            }
                        }

                        // OBTENER CLIENTE/VENTAS AL CREDITO
                        if (filaActual == 15 || leyendoClientes)
                        {
                            var valorNombre = reader.GetValue(0);
                            var valorMonto = reader.GetValue(6);

                            if (valorNombre != null && !string.IsNullOrWhiteSpace(valorNombre.ToString()))
                            {
                                string nombreLimpio = valorNombre.ToString().Replace("  ", " ").Trim();

                                decimal montoActual = 0;
                                decimal.TryParse(valorMonto?.ToString(), out montoActual);

                                if (clientesAgrupados.ContainsKey(nombreLimpio))
                                {
                                    clientesAgrupados[nombreLimpio] += montoActual;
                                }
                                else
                                {
                                    clientesAgrupados.Add(nombreLimpio, montoActual);
                                }
                                leyendoClientes = true;
                            }
                            else
                            {
                                if (flagclientecredito == 1)
                                {
                                    leyendoClientes = false;
                                }
                                flagclientecredito = 1;
                            }
                        }

                        // RESUMEN DE VARIACIONES
                        if (FilasVariaciones.Contains(filaActual))
                        {
                            var Varia_combus_nombre = reader.GetValue(16);
                            var Varia_combus_monto = reader.GetValue(18);
                            decimal montoActualVariacion = 0;
                            decimal.TryParse(Varia_combus_monto?.ToString(), out montoActualVariacion);

                            var Varia_combus_nombre_corregido = Varia_combus_nombre?.ToString()?.Replace(" ", "");

                            if (Varia_combus_nombre_corregido == "GLP")
                            {
                                registro.DescuentoGLP = Varia_combus_monto ?? 0;
                            }
                            else
                            {
                                DescuentoLiquidos_Total = DescuentoLiquidos_Total + montoActualVariacion;
                            }
                        }

                        // RESUMEN DE HERMES
                        if (filaActual >= 60 && filaActual <= 130)
                        {
                            var celdaIdentificadora = reader.GetValue(14);
                            string textoCelda = celdaIdentificadora?.ToString()?.Trim() ?? "";

                            // 4. Búsqueda Segura de la Cabecera de Hermes
                            if (!HermesTablaEncontra && textoCelda.Contains(HermesPalabraClaveCabecera, StringComparison.OrdinalIgnoreCase))
                            {
                                HermesTablaEncontra = true;
                                HermesLeyendoTablaFlotante = true;

                                filaActual++;
                                continue;
                            }

                            // Fase B: Extraer data de Hermes
                            if (HermesLeyendoTablaFlotante)
                            {
                                if (!string.IsNullOrWhiteSpace(textoCelda))
                                {
                                    var valorTabla = reader.GetValue(14);
                                    decimal montoActualHermes = 0;
                                    decimal.TryParse(valorTabla?.ToString(), out montoActualHermes);

                                    switch (ContadorHermes)
                                    {
                                        case 1: registro.Hermes_monto_liquido = montoActualHermes; break;
                                        case 2: registro.Hermes_monto_GLP = montoActualHermes; break;
                                        case 3: registro.Hermes_monto_GNV1 = montoActualHermes; break;
                                        case 4: registro.Hermes_monto_GNV2 = montoActualHermes; break;
                                    }
                                    ContadorHermes++;
                                }
                                else
                                {
                                    HermesLeyendoTablaFlotante = false;
                                }
                            }
                        }

                        filaActual++;
                    }

                    foreach (var entrada in clientesAgrupados)
                    {
                        registro.AgregarClienteCredito(entrada.Key, entrada.Value);
                    }
                    registro.DescuentoLiquidos = DescuentoLiquidos_Total;

                    listaResultados.Add(registro);

                } while (reader.NextResult());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en {Path.GetFileName(ruta)}: {ex.Message}");
            }
        });

        string json = JsonSerializer.Serialize(listaResultados,
            new JsonSerializerOptions { WriteIndented = true });

        Console.WriteLine(json);
    }
}