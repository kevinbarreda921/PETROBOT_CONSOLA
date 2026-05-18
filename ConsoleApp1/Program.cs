using ConsoleApp1.Services;
using System;
using System.Diagnostics;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch timerGlobal = Stopwatch.StartNew();
            LoggerService.LimpiarLogs();

            try
            {
                var consolidador = new ConsolidadorVentasService();
                consolidador.Procesar();
            }
            catch (Exception ex)
            {
                LoggerService.Error("SISTEMA", "Main", $"Ocurrió un error fatal: {ex.Message}");
                Console.WriteLine($"[x] Ocurrió un error fatal:\n{ex.ToString()}");
            }
            finally
            {
                string rutaLog = System.IO.Path.GetFullPath(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\", "reporte_proceso.json"));
                LoggerService.GuardarJson(rutaLog);
            }

            timerGlobal.Stop();
            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));
        }
    }
}
