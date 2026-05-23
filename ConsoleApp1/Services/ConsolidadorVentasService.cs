using OfficeOpenXml;
using ConsoleApp1.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApp1.Services
{
    public class ConsolidadorVentasService
    {
        private readonly LectorExcelService _lector;
        private readonly EscritorExcelService _escritor;

        public ConsolidadorVentasService()
        {
            _lector = new LectorExcelService();
            _escritor = new EscritorExcelService();
        }

        public void Procesar()
        {
            var CONTADOR_dES = 1;
            Console.WriteLine("Iniciando lectura de archivos...");
            var listaGrifosProcesar = _lector.LeerPartesDiarios();


            string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
            string carpetaRegistroVentas = Path.Combine(rutaProyecto, "ArchivosExcel", "Registro_ventas");

            var archivosExcel = Directory.GetFiles(carpetaRegistroVentas, "*.xlsx");
            if (archivosExcel.Length == 0)
            {
                Console.WriteLine("[x] No se encontró ningún archivo Excel maestro en la carpeta Registro_ventas.");
                return;
            }
            string rutaExcel = archivosExcel[0];

            // Eliminamos la pasada inicial redundante de ExcelDataReader


            Console.WriteLine("Abriendo archivo Excel maestro con EPPlus...");
            ExcelPackage.License.SetNonCommercialPersonal("PETROBOT");
            using var package = new ExcelPackage(new FileInfo(rutaExcel));
            package.Compression = CompressionLevel.BestSpeed; // Reducir overhead de CPU al comprimir
            var workbook = package.Workbook;

            foreach (var archivoGrifoActual in listaGrifosProcesar)
            {
                string? grifoObjetivo = archivoGrifoActual.Grifo;
                if (string.IsNullOrEmpty(grifoObjetivo)) continue;

                LoggerService.Info(grifoObjetivo, archivoGrifoActual.Archivo ?? "", $" CONTADOR {CONTADOR_dES}");
                CONTADOR_dES++;

                // Buscamos la hoja cuyo nombre contenga el grifoObjetivo de forma flexible (ignora mayúsculas)
                var hojaEPPlus = workbook.Worksheets.FirstOrDefault(w => w.Name.IndexOf(grifoObjetivo, StringComparison.OrdinalIgnoreCase) >= 0);
                
                if (hojaEPPlus == null)
                {
                    LoggerService.Error(grifoObjetivo, "MAESTRO", $"No se encontró ninguna hoja para el grifo {grifoObjetivo} en EPPlus.");
                    continue;
                }

                // LoggerService.Info(grifoObjetivo, archivoGrifoActual.Archivo, $"Hoja encontrada y lista para procesar.");
                
                // Mapeamos las filas de fechas directamente desde memoria (sin leer el archivo de nuevo)
                var mapaFechasFilas = _escritor.MapearFechasHoja(hojaEPPlus);

                    var fechasDelGrifo = archivoGrifoActual.ListVenta
                        .Where(v => !string.IsNullOrEmpty(v.Dia))
                        .Select(v => v.Dia!)
                        .Distinct()
                        .ToList();

                    // Obtener configuración global
                    ConfiguracionService.ConfigGlobal.Grifos.TryGetValue(grifoObjetivo, out var configGrifoRoot);
                    var configColumnas = configGrifoRoot?.Escritura;
                    var configClientes = configGrifoRoot?.FilasClientesCreditos;

                    // Invertir diccionario de clientes una sola vez por Grifo
                    var clienteAColumna = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    if (configClientes != null)
                    {
                        foreach (var kvp in configClientes)
                        {
                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                            {
                                clienteAColumna[kvp.Value.Trim()] = kvp.Key.Trim();
                            }
                        }
                    }

            
                foreach (string fechaABuscar in fechasDelGrifo)
                    {
                        if (mapaFechasFilas.TryGetValue(fechaABuscar, out int filaDestino))
                        {
                        LoggerService.Info(grifoObjetivo, archivoGrifoActual.Archivo, $" El grifo {grifoObjetivo} del dia {fechaABuscar} procesado correctamente");

                   
                            if (configColumnas != null)
                            {
                                var ventaParaEscribir = archivoGrifoActual.ListVenta.FirstOrDefault(v => v.Dia == fechaABuscar);
                                if (ventaParaEscribir != null)
                                {
                                    _escritor.EscribirFila(hojaEPPlus, ventaParaEscribir, filaDestino, configColumnas.Columnas);
                                    
                                    if (clienteAColumna.Count > 0)
                                    {
                                        _escritor.EscribirClientesCredito(hojaEPPlus, ventaParaEscribir, filaDestino, clienteAColumna, grifoObjetivo, archivoGrifoActual.Archivo ?? "DESCONOCIDO");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"[!] No se encontró mapeo de escritura para el grifo: {grifoObjetivo}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"[x] La fecha {fechaABuscar} NO EXISTE en el Maestro para el grifo {grifoObjetivo}");
                        }
                    }
            }

            Console.WriteLine("Guardando archivo Excel...");
            package.Save();
            Console.WriteLine($"[✓] Archivo Excel guardado correctamente.");
        }
    }
}
