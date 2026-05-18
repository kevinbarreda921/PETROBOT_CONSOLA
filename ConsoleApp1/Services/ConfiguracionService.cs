using ConsoleApp1.Models.Configuracion;
using System;
using System.IO;
using System.Text.Json;

namespace ConsoleApp1.Services
{
    public static class ConfiguracionService
    {
        public static ConfigRoot ConfigGlobal { get; private set; } = new();

        static ConfiguracionService()
        {
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
                    ConfigGlobal = JsonSerializer.Deserialize<ConfigRoot>(jsonTexto) ?? new ConfigRoot();
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
    }
}
