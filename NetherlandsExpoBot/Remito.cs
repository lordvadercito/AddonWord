using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace NetherlandsExpoBot
{
    public class Remito
    {
        public List<RemitoItem> remitoList { get; set; }

        public Remito GetRemitoNro(int remitoValue)
        {
            List<RemitoItem> remitoList = new List<RemitoItem>();
            var connStr = "Server = 10.1.100.26,50624; Database = sanitariosDB; User Id = gestorDevesa; Password = juampa;";//Production connection
            string storeProcedure = "PD_Informe_Sanitarios_Exportacion";
            using (var conn = new SqlConnection(connStr))
            using (var cmd = new SqlCommand(storeProcedure, conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@remito", remitoValue);
                conn.Open();
                System.Diagnostics.Debug.WriteLine(cmd.CommandText);

                SqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    RemitoItem rmi = new RemitoItem();
                    rmi.RemitoID = (string)rdr["RemitoID"];
                    rmi.Cajas = (int)rdr["Cajas"];
                    rmi.Peso = (int)rdr["Peso"];
                    rmi.Bruto = (int)rdr["Bruto"];
                    rmi.Descripcion = (string)rdr["Descripcion"];
                    rmi.Observaciones = (string)rdr["Observaciones"];
                    rmi.Precinto = (string)rdr["Precinto"];
                    rmi.Contenedor = (string)rdr["Contenedor"];
                    rmi.Destino = (string)rdr["Destino"];
                    rmi.Domicilio = (string)rdr["Domicilio"];
                    rmi.Vapor = (string)rdr["Vapor"];
                    remitoList.Add(rmi);
                }
            }

            Remito remitos = new Remito();
            remitos.remitoList = remitoList;
            return remitos;
        }
    }
}