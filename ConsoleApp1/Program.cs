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

            try
            {
                var consolidador = new ConsolidadorVentasService();
                consolidador.Procesar();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[x] Ocurrió un error fatal:\n{ex.ToString()}");
            }

            timerGlobal.Stop();
            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));
        }
    }
}
