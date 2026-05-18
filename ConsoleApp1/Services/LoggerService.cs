using ConsoleApp1.Models.Log;
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text.Json;

namespace ConsoleApp1.Services
{
    public static class LoggerService
    {
        private static readonly ConcurrentDictionary<string, LogProcesoGrifo> _logs = new();
        private static readonly object _fileLock = new object();

        public static void LimpiarLogs()
        {
            _logs.Clear();
        }

        public static void Info(string grifo, string archivo, string descripcion)
        {
            RegistrarMensaje(grifo, archivo, "INFO", descripcion);
            Console.WriteLine($"[INFO] {descripcion}");
        }

        public static void Error(string grifo, string archivo, string descripcion)
        {
            RegistrarMensaje(grifo, archivo, "ERROR", descripcion);
            Console.WriteLine($"[ERROR] {descripcion}");
        }

        private static void RegistrarMensaje(string grifo, string archivo, string estado, string descripcion)
        {
            // Usamos una clave combinada de Grifo + Archivo en caso de que un grifo tenga varios archivos
            string key = $"{grifo}_{archivo}";

            var logGrifo = _logs.GetOrAdd(key, k => new LogProcesoGrifo { Grifo = grifo, Archivo = archivo });
            
            lock (logGrifo.Mensajes)
            {
                logGrifo.Mensajes.Add(new LogDetalle { Estado = estado, Descripcion = descripcion });
            }
        }

        public static void GuardarJson(string rutaArchivo)
        {
            try
            {
                lock (_fileLock)
                {
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    string json = JsonSerializer.Serialize(_logs.Values, options);
                    File.WriteAllText(rutaArchivo, json);
                    Console.WriteLine($"\n[✓] Log guardado correctamente en: {rutaArchivo}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[x] Error fatal guardando el archivo de log: {ex.Message}");
            }
        }
    }
}
