using ClosedXML.Excel;
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
            Console.WriteLine("Iniciando lectura de archivos...");
            var listaGrifosProcesar = _lector.LeerPartesDiarios();

            List<string> misGrifos = listaGrifosProcesar
                .Where(g => !string.IsNullOrEmpty(g.Grifo))
                .Select(g => g.Grifo!)
                .Distinct()
                .ToList();

            string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
            string carpetaRegistroVentas = Path.Combine(rutaProyecto, "ArchivosExcel", "Registro_ventas");

            var archivosExcel = Directory.GetFiles(carpetaRegistroVentas, "*.xlsx");
            if (archivosExcel.Length == 0)
            {
                Console.WriteLine("[x] No se encontró ningún archivo Excel maestro en la carpeta Registro_ventas.");
                return;
            }
            string rutaExcel = archivosExcel[0];

            Console.WriteLine("Mapeando estructura del archivo Excel maestro...");
            List<HojaGrifoMapeada> listaMapeada = _escritor.MapearEstructuraMaestro(rutaExcel, misGrifos);

            var diccionarioGrifos = listaMapeada.ToDictionary(
                g => g.Grifo ?? "SIN_NOMBRE",
                g => g,
                StringComparer.OrdinalIgnoreCase
            );

            Console.WriteLine("Abriendo archivo Excel maestro...");
            using var workbook = new XLWorkbook(rutaExcel);

            foreach (string grifoObjetivo in misGrifos)
            {
                if (diccionarioGrifos.TryGetValue(grifoObjetivo, out var miGrifo))
                {
                    Console.WriteLine($"[✓] Grifo encontrado instantáneamente en la hoja: {miGrifo.Hoja}");

                    if (!workbook.TryGetWorksheet(miGrifo.Hoja, out var hojaClosedXML))
                    {
                        Console.WriteLine($"[x] No se pudo abrir la hoja {miGrifo.Hoja} con ClosedXML.");
                        continue;
                    }

                    var archivoGrifoActual = listaGrifosProcesar.FirstOrDefault(g => g.Grifo == grifoObjetivo);
                    if (archivoGrifoActual == null) continue;

                    var fechasDelGrifo = archivoGrifoActual.ListVenta
                        .Where(v => !string.IsNullOrEmpty(v.Dia))
                        .Select(v => v.Dia!)
                        .Distinct()
                        .ToList();

                    // Obtener configuración global
                    ConfiguracionService.ConfigGlobal.Grifos.TryGetValue(grifoObjetivo, out var configGrifoRoot);
                    var configColumnas = configGrifoRoot?.Escritura;
                    var configClientes = configGrifoRoot?.FilasClientesCreditos;

                    foreach (string fechaABuscar in fechasDelGrifo)
                    {
                        if (miGrifo.MapaFechasFilas.TryGetValue(fechaABuscar, out int filaDestino))
                        {
                            Console.WriteLine($"[✓] Escribiendo datos de la fecha {fechaABuscar} en la FILA: {filaDestino}");

                            if (configColumnas != null)
                            {
                                var ventaParaEscribir = archivoGrifoActual.ListVenta.FirstOrDefault(v => v.Dia == fechaABuscar);
                                if (ventaParaEscribir != null)
                                {
                                    _escritor.EscribirFila(hojaClosedXML, ventaParaEscribir, filaDestino, configColumnas.Columnas);
                                    
                                    if (configClientes != null && configClientes.Count > 0)
                                    {
                                        _escritor.EscribirClientesCredito(hojaClosedXML, ventaParaEscribir, filaDestino, configClientes, grifoObjetivo);
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
            }

            Console.WriteLine("Guardando archivo Excel...");
            workbook.Save();
            Console.WriteLine($"[✓] Archivo Excel guardado correctamente.");
        }
    }
}
