using ConsoleApp1;
using System.Diagnostics;
using System.Text.Json;


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
foreach (string auto in misGrifos)
{
    string grifoObjetivo = auto;

    if (diccionarioGrifos.TryGetValue(grifoObjetivo, out var miGrifo))
    {
        Console.WriteLine($"[✓] Grifo encontrado instantáneamente en la hoja: {miGrifo.Hoja}");

        var fechasDelGrifo = listaGrifosProcesar
            .Where(g => g.Grifo == grifoObjetivo)
            .SelectMany(g => g.ListVenta)
            .Where(v => !string.IsNullOrEmpty(v.Dia))
            .Select(v => v.Dia!)
            .Distinct()
            .ToList();

        foreach (string mFECHAS in fechasDelGrifo)
        {
            string fechaABuscar = mFECHAS;

            if (miGrifo.MapaFechasFilas.TryGetValue(fechaABuscar, out int filaDestino))
            {
                Console.WriteLine($"[✓] La fecha {fechaABuscar} está en la FILA: {filaDestino}");
            }
            else
            {
                Console.WriteLine($"[x] La fecha {fechaABuscar} NO EXISTE");
            }
        }

    }
}
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

