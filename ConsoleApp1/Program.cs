using ConsoleApp1;
using System.Diagnostics;
using System.Text.Json;
using ClosedXML.Excel;
using System.IO;


Stopwatch timerGlobal = Stopwatch.StartNew();

var listaGrifosProcesar = excel_data_reader_parallel.f_excel_data_reader_parallel();

var escritorExcel = new ExcelDataWrite();

List<string> misGrifos = listaGrifosProcesar
    .Where(g => !string.IsNullOrEmpty(g.Grifo))
    .Select(g => g.Grifo!)
    .Distinct()
    .ToList();

List<string> misFECHAS = listaGrifosProcesar
    .SelectMany(g => g.ListVenta)
    .Where(v => !string.IsNullOrEmpty(v.Dia))
    .Select(v => v.Dia!)
    .Distinct()
    .ToList();

List<ExcelDataWrite.HojaGrifoMapeada> listaMapeada = escritorExcel.MapearEstructuraMaestro(misGrifos);

var diccionarioGrifos = listaMapeada.ToDictionary(
    g => g.Grifo ?? "SIN_NOMBRE",
    g => g,
    StringComparer.OrdinalIgnoreCase
);
string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
string rutaExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "Registro_ventas", "REGISTRO VENTAS -  2026- 01.xlsx");

Console.WriteLine("Abriendo archivo Excel maestro...");
using var workbook = new XLWorkbook(rutaExcel);

foreach (string auto in misGrifos)
{
    string grifoObjetivo = auto;

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

        // Obtener configuración de columnas para este grifo
        excel_data_reader_parallel.ConfigGlobal.Grifos.TryGetValue(grifoObjetivo, out var configGrifoRoot);
        var configColumnas = configGrifoRoot?.Escritura;
        var configClientes = configGrifoRoot?.FilasClientesCreditos;

        foreach (string mFECHAS in fechasDelGrifo)
        {
            string fechaABuscar = mFECHAS;

            if (miGrifo.MapaFechasFilas.TryGetValue(fechaABuscar, out int filaDestino))
            {
                Console.WriteLine($"[✓] Escribiendo datos de la fecha {fechaABuscar} en la FILA: {filaDestino}");

                if (configColumnas != null)
                {
                    var ventaParaEscribir = archivoGrifoActual.ListVenta.FirstOrDefault(v => v.Dia == fechaABuscar);
                    if (ventaParaEscribir != null)
                    {
                        escritorExcel.EscribirFila(hojaClosedXML, ventaParaEscribir, filaDestino, configColumnas.Columnas);
                        
                        if (configClientes != null && configClientes.Count > 0)
                        {
                            escritorExcel.EscribirClientesCredito(hojaClosedXML, ventaParaEscribir, filaDestino, configClientes, grifoObjetivo);
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
                Console.WriteLine($"[x] La fecha {fechaABuscar} NO EXISTE en el Maestro");
            }
        }
    }
}

Console.WriteLine("Guardando archivo Excel...");
workbook.Save();
Console.WriteLine($"[✓] Archivo Excel guardado correctamente.");
timerGlobal.Stop();

Console.WriteLine("\n" + new string('=', 30));
Console.WriteLine($"PROCESO FINALIZADO");
Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
Console.WriteLine(new string('=', 30));


//// 2. Configurar la serialización para que sea legible (Indented)
//var opcionesJson = new JsonSerializerOptions
//{
//    WriteIndented = true // Esto le da formato de escalera al texto para que sea fácil de leer
//};

//// 3. Convertir la lista a una cadena JSON
//string jsonResultado = JsonSerializer.Serialize(estructuraMapeada, opcionesJson);

//// 4. Imprimir en la consola
//Console.WriteLine("=== ESTRUCTURA DE EXCEL MAESTRO MAPEADA ===");
//Console.WriteLine(jsonResultado);

