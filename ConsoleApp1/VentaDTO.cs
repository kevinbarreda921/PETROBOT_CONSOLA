using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1
{

    public  class VentaDTO
    {

        public string? Hoja { get; set; }
        public string? Dia { get; set; }
        public object? Venta_GPL { get; set; }
        public object? Venta_GNV { get; set; }
        public object? Total_venta_acumulada { get; set; }
        public object? Total_Tarjeta_de_Credito_Liquidos { get; set; }
        public object? Total_Tarjeta_de_Credito_GLP { get; set; }
        public object? Total_Tarjeta_de_Credito_GNV { get; set; }
        public object? ErrorMaquina { get; set; }
        public object? Recaudo_Cofide_GNV { get; set; }
        public object? Gastos { get; set; }
        public object? Ventas_con_transferencia { get; set; }
        public object? DescuentoLiquidos { get; set; }
        public object? DescuentoGLP { get; set; }

        public object? Hermes_monto_liquido { get; set; }
        public object? Hermes_monto_GLP { get; set; }
        public object? Hermes_monto_GNV1 { get; set; }
        public object? Hermes_monto_GNV2 { get; set; }

        // Lista
        public List<ClienteCreditoDTO> ListClienteCredito { get; set; }

        // Constructor vacío
        public VentaDTO()
        {
            ListClienteCredito = new List<ClienteCreditoDTO>();
        }

        // Constructor con parámetros
        public VentaDTO(
            string? hoja = null,
            object? venta_GPL = null,
            object? venta_GNV = null,
            object? total_venta_acumulada = null,
            object? total_Tarjeta_de_Credito_Liquidos = null,
            object? total_Tarjeta_de_Credito_GLP = null,
            object? total_Tarjeta_de_Credito_GNV = null,
            object? errorMaquina = null,
            object? recaudo_Cofide_GNV = null,
            object? gastos = null,
            object? ventas_con_transferencia = null,
            decimal descuentoLiquidos = 0,
            decimal descuentoGLP = 0,
            decimal hermes_monto_liquido = 0,
            decimal hermes_monto_GLP = 0,
            decimal hermes_monto_GNV1 = 0,
            decimal hermes_monto_GNV2 = 0
            )
        {
            Hoja = hoja;
            Venta_GPL = venta_GPL;
            Venta_GNV = venta_GNV;
            Total_venta_acumulada = total_venta_acumulada;
            Total_Tarjeta_de_Credito_Liquidos = total_Tarjeta_de_Credito_Liquidos;
            Total_Tarjeta_de_Credito_GLP = total_Tarjeta_de_Credito_GLP;
            Total_Tarjeta_de_Credito_GNV = total_Tarjeta_de_Credito_GNV;
            ErrorMaquina = errorMaquina;
            Recaudo_Cofide_GNV = recaudo_Cofide_GNV;
            Gastos = gastos;
            Ventas_con_transferencia = ventas_con_transferencia;
            DescuentoLiquidos = descuentoLiquidos;
            DescuentoGLP = descuentoGLP;
            Hermes_monto_liquido = hermes_monto_liquido;
            Hermes_monto_GLP = hermes_monto_GLP;
            Hermes_monto_GNV1 = hermes_monto_GNV1;
            Hermes_monto_GNV2 = hermes_monto_GNV2;
            ListClienteCredito = new List<ClienteCreditoDTO>();
        }
        public void AgregarClienteCredito(string cliente, object? monto)
        {
            ListClienteCredito.Add(new ClienteCreditoDTO(cliente, monto));
        }
    }

    public class ClienteCreditoDTO
    {
        public string? Cliente { get; set; }
        public object? Monto { get; set; }

        public ClienteCreditoDTO(string? cliente, object? monto)
        {
            Cliente = cliente;
            Monto = monto;
        }
    }

}
