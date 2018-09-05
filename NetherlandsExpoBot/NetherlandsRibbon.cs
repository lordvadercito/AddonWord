using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;

namespace NetherlandsExpoBot
{
    public partial class NetherlandsRibbon
    {
        private void NetherlandsRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void remitoSearch_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string remito = this.remitoSearch.Text;
            //int remitoInt = Int32.Parse(remito);
            Remito remitoBusqueda = new Remito();
            Remito remitoResult = remitoBusqueda.GetRemitoNro(remito);
            int countRemitos = remitoResult.remitoList.Count();

            // Elementos de primera parte del anexo
            Globals.ThisDocument.rtcNombreDestino.Text = remitoResult.remitoList[0].Destino.ToString();
            Globals.ThisDocument.rtcDireccionDestino.Text = remitoResult.remitoList[0].Domicilio.ToString();
            Globals.ThisDocument.rtcPrecinto.Text = remitoResult.remitoList[0].Precinto.ToString();
            Globals.ThisDocument.rtcContenedor.Text = remitoResult.remitoList[0].Contenedor.ToString();
            Globals.ThisDocument.rtcVapor.Text = remitoResult.remitoList[0].Vapor.ToString();

            //Tabla de productos y elementos acumulativos
            string mercaderiasNombre = "rtcMercaderiaGrid";
            string bultosNombre = "rtcBultosGrid";
            string netoNombre = "rtcPesoNetoGrid";
            string mataderoNombre = "rtcMataderoGrid";
            string fabricaNombre = "rtcFabricaGrid";
            string frigorificoNombre = "rtcFrigorificoGrid";
            int sumaBultos = 0;
            int sumaBruto = 0;
            int sumaNeto = 0;

            for (int i = 0; i < countRemitos; i++)
            {
                string mercaderiaPosicion = string.Concat(mercaderiasNombre, i);
                string bultosPosicon = string.Concat(bultosNombre, i);
                string netoPosicion = string.Concat(netoNombre, i);
                string mataderoPosicion = string.Concat(mataderoNombre, i);
                string fabricaPosicion = string.Concat(fabricaNombre, i);
                string frigorificoPosicion = string.Concat(frigorificoNombre, i);

                RichTextContentControl mercaderiaPosicionada = Globals.ThisDocument.hashTableElementos.GetControl(mercaderiaPosicion.ToString());
                RichTextContentControl netoPosicionado = Globals.ThisDocument.hashTableElementos.GetControl(netoPosicion.ToString());
                RichTextContentControl bultosPosicionados = Globals.ThisDocument.hashTableElementos.GetControl(bultosPosicon.ToString());
                RichTextContentControl mataderoPosicionado = Globals.ThisDocument.hashTableElementos.GetControl(mataderoPosicion.ToString());
                RichTextContentControl fabricaPosicionada = Globals.ThisDocument.hashTableElementos.GetControl(fabricaPosicion.ToString());
                RichTextContentControl frigorificoPosicionado = Globals.ThisDocument.hashTableElementos.GetControl(frigorificoPosicion.ToString());

                mercaderiaPosicionada.Text = remitoResult.remitoList[i].Descripcion.ToString();
                bultosPosicionados.Text = remitoResult.remitoList[i].Cajas.ToString();
                netoPosicionado.Text = remitoResult.remitoList[i].Peso.ToString();
                mataderoPosicionado.Text = "5039";
                fabricaPosicionada.Text = "5039";
                frigorificoPosicionado.Text = "5039";

                sumaBruto = sumaBruto + int.Parse(remitoResult.remitoList[i].Bruto.ToString());
                sumaNeto = sumaNeto + int.Parse(remitoResult.remitoList[i].Peso.ToString());
                sumaBultos = sumaBultos + int.Parse(remitoResult.remitoList[i].Cajas.ToString());

            }

            RichTextContentControl netoTotalPosicionado = Globals.ThisDocument.hashTableElementos.GetControl("rtcTotalesKilos"); 
            RichTextContentControl bultosTotalPosicionado = Globals.ThisDocument.hashTableElementos.GetControl("rtcTotalBultos");
            RichTextContentControl netoGrid = Globals.ThisDocument.hashTableElementos.GetControl("rtcPesoNetoGrid");
            RichTextContentControl bultosGrid = Globals.ThisDocument.hashTableElementos.GetControl("rtcBultosGrid");
            RichTextContentControl bultosSup = Globals.ThisDocument.hashTableElementos.GetControl("rtcBultos");
            RichTextContentControl netoSup = Globals.ThisDocument.hashTableElementos.GetControl("rtcPesoNeto");
            RichTextContentControl brutoSup = Globals.ThisDocument.hashTableElementos.GetControl("rtcPesoBruto");
            RichTextContentControl cuotaSpanish = Globals.ThisDocument.hashTableElementos.GetControl("rtcCuotaSpanish");
            RichTextContentControl cuotaEnglish = Globals.ThisDocument.hashTableElementos.GetControl("rtcCuotaEnglish");

            netoTotalPosicionado.Text = sumaNeto.ToString();
            bultosTotalPosicionado.Text = sumaBultos.ToString();
            bultosGrid.Text = sumaBultos.ToString();
            netoGrid.Text = sumaNeto.ToString();
            bultosSup.Text = sumaBultos.ToString();
            netoSup.Text = sumaNeto.ToString();
            brutoSup.Text = sumaBruto.ToString();

            String primerProducto = remitoResult.remitoList[0].Descripcion.ToString();

            if (primerProducto.Contains("(W)"))
            {
                cuotaSpanish.Text = "481";
                cuotaEnglish.Text = "481";
            }
            else if (primerProducto.Contains("(H)"))
            {
                cuotaSpanish.Text = "HILTON";
                cuotaEnglish.Text = "HILTON";
            }

        }
    }
}
