using ERPConnect;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace BPL
{
    public class SAPHelp
    {
        public static DataTable GetMaterials()
        {
            DataTable resulttable  = new DataTable();
            using (R3Connection con = RfcConnection.GetR3PConnection())
            {
                con.Open(false);
                ERPConnect.Utils.ReadTable table = new ERPConnect.Utils.ReadTable(con);

                table.AddField("MATNR");
                table.AddField("MAKTX");
                table.AddCriteria("SPRAS = 'DE'");
                table.TableName = "MAKT";
                table.RowCount = 10;

                table.Run();

                resulttable = table.Result;

                for (int i = 0; i < resulttable.Rows.Count; i++)
                {
                    Console.WriteLine(
                        resulttable.Rows[i]["MATNR"].ToString() + " " +
                        resulttable.Rows[i]["MAKTX"].ToString());
                }

                return resulttable;
            }            
        }
    }
}
