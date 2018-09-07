using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RaportyCoC
{
    class SqlOpeations
    {
        public static List<string> GetBoxesIdFromPalletId(string[] pallets)
        {
            HashSet<string> result = new HashSet<string>();
            DataTable sqlTable = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = @"Data Source=MSTMS010;Initial Catalog=MES;User Id=mes;Password=mes;";

            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = String.Format(@"SELECT distinct Box_LOT_NO,Palet_LOT_NO FROM tb_WyrobLG_opakowanie WHERE ");

            for (int i = 0; i < pallets.Length; i++)
            {
                if (i == 0)
                {
                    command.CommandText += "Palet_LOT_NO = @" + i + "box";
                    command.Parameters.AddWithValue("@" + i + "box", pallets[i]);
                }
                else
                {
                    command.CommandText += " OR Palet_LOT_NO = @" + i + "box";
                    command.Parameters.AddWithValue("@" + i + "box", pallets[i]);
                }
            }

            command.CommandText += ";";
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            try
            {
                adapter.Fill(sqlTable);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.HResult);
            }
            foreach (DataRow row in sqlTable.Rows)
            {
                result.Add(row["Box_LOT_NO"].ToString());
            }
            return result.ToList();
        }

        public static Dictionary<string, PcbTesterMeasurements> GetMeasurementsForBoxes(string[] boxesId)
        {
            DataTable sqlTable = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = @"Data Source=MSTMS010;Initial Catalog=MES;User Id=mes;Password=mes;";

            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = String.Format(@"SELECT serial_no,inspection_time,wip_entity_name,sdcm,cct,lm,lm_w,cri,cct,v,x,y,Box_LOT_NO,NC12,result,CUSTOMER_ORDER_NO FROM v_tester_measurements_all WHERE (");

            for (int i = 0; i < boxesId.Length; i++)
            {
                if (i == 0)
                {
                    command.CommandText += "Box_LOT_NO = @" + i + "box";
                    command.Parameters.AddWithValue("@" + i + "box", boxesId[i]);
                }
                else
                {
                    command.CommandText += " OR Box_LOT_NO = @" + i + "box";
                    command.Parameters.AddWithValue("@" + i + "box", boxesId[i]);
                }
            }

            command.CommandText += ") AND result='OK' ORDER BY inspection_time DESC;";
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(sqlTable);

            return dataTableToDict(sqlTable);
        }

        public static Dictionary<string, PcbTesterMeasurements> GetMeasurementsForOrderNo(List<string> orderNumbers)
        {
            DataTable sqlTable = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = @"Data Source=MSTMS010;Initial Catalog=MES;User Id=mes;Password=mes;";

            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = String.Format(@"SELECT serial_no,Wysylki_Data,inspection_time,wip_entity_name,sdcm,cct,lm,lm_w,cri,cct,v,x,y,Box_LOT_NO,NC12,result,CUSTOMER_ORDER_NO FROM v_tester_measurements_all WHERE (");
            for (int i=0;i<orderNumbers.Count;i++)
            {
                if (i==0)
                {
                    command.CommandText += "CUSTOMER_ORDER_NO=@order" + i;
                    command.Parameters.AddWithValue("@order"+i, orderNumbers[i]);
                }
                else
                {
                    command.CommandText += " OR CUSTOMER_ORDER_NO=@order" + i + " ";
                    command.Parameters.AddWithValue("@order" + i, orderNumbers[i]);
                }
            }
            command.CommandText+=") AND result = 'OK' ORDER BY inspection_time DESC; ";
            
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(sqlTable);

            HashSet<string> datesOfShipping = new HashSet<string>();
            foreach (DataRow     row     in sqlTable.Rows)
            {
                datesOfShipping.Add(row["Wysylki_Data"].ToString());
            }

            List<string> choosenDates = datesOfShipping.ToList();
            
            if (datesOfShipping.Count>0)
            {
                using (chooseDateOfShipment chooseForm = new chooseDateOfShipment(datesOfShipping.ToList()))
                {
                    if (chooseForm.ShowDialog() == DialogResult.OK)
                    {
                        choosenDates.Clear();
                        choosenDates = chooseForm.choosenDates;
                    }

                }

                DataTable tempTable = sqlTable.Clone();
                foreach (DataRow row in sqlTable.Rows)
                {
                    if (choosenDates.Contains( row["Wysylki_Data"].ToString()))
                    {
                        tempTable.Rows.Add(row.ItemArray);
                    }
                }

                sqlTable = tempTable;
            }
            
            

            return dataTableToDict(sqlTable);
        }

        public static Dictionary<string, PcbTesterMeasurements> GetMeasurementsForPallets(string palletsID)
        {
            DataTable sqlTable = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = @"Data Source=MSTMS010;Initial Catalog=MES;User Id=mes;Password=mes;";

            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = String.Format(@"SELECT serial_no,inspection_time,wip_entity_name,sdcm,cct,lm,lm_w,cri,cct,v,x,y,Box_LOT_NO,Palet_LOT_NO,NC12,result,CUSTOMER_ORDER_NO FROM v_tester_measurements_all WHERE (");

            for (int i = 0; i < palletsID.Length; i++)
            {
                if (i == 0)
                {
                    command.CommandText += "Palet_LOT_NO = " + i + "@box";
                    command.Parameters.AddWithValue(i + "@box", palletsID[i]);
                }
                else
                {
                    command.CommandText += "OR Palet_LOT_NO = " + i + "@box";
                    command.Parameters.AddWithValue(i + "@box", palletsID[i]);
                }
            }

            command.CommandText += ") AND result='OK' ORDER BY inspection_time DESC;";
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(sqlTable);

            return dataTableToDict(sqlTable);
        }

        private static bool IsInt(string s)
        {
            int x = 0;
            return int.TryParse(s, out x);
        }

        private static Dictionary<string, PcbTesterMeasurements> dataTableToDict(DataTable sqlTable)
        {
            Dictionary<string, PcbTesterMeasurements> result = new Dictionary<string, PcbTesterMeasurements>();
            if (sqlTable.Rows.Count > 0)
            {
                Dictionary<string, string> lotToModel = new Dictionary<string, string>();

                foreach (DataRow row in sqlTable.Rows)
                {
                    string serial = row["serial_no"].ToString();
                    if (result.ContainsKey(serial))
                        continue;

                    string lot = row["wip_entity_name"].ToString();
                    string model = "";

                    if (!IsInt(lot))
                        continue;
                    string orderNo = row["CUSTOMER_ORDER_NO"].ToString();

                    model = row["NC12"].ToString();
                    if (!lotToModel.ContainsKey(lot))
                    {
                        lotToModel.Add(lot, model);
                    }
                    string cxString = row["x"].ToString().Replace(".", ",");
                    string cyString = row["y"].ToString().Replace(".", ",");
                    if (cxString == "" || cyString == "")
                        continue;

                    if (model == "")
                        continue;
                    double cx = Convert.ToDouble(cxString, new CultureInfo("pl-PL"));
                    double cy = double.Parse(row["y"].ToString().Replace(".", ","));
                    double sdcm = double.Parse(row["sdcm"].ToString());
                    double cct = double.Parse(row["cct"].ToString());
                    double vf = double.Parse(row["v"].ToString());
                    double lm = double.Parse(row["lm"].ToString());
                    double lmW = double.Parse(row["lm_w"].ToString());
                    double cri = double.Parse(row["cri"].ToString());

                    DateTime inspTime = DateTime.Parse(row["inspection_time"].ToString());

                    PcbTesterMeasurements newPcb = new PcbTesterMeasurements(cx, cy, sdcm, cct, inspTime, model, vf, lm, lmW, cri, cct, orderNo);
                    result.Add(serial, newPcb);
                }

            }
            return result;
        }
        public static Dictionary<string, ModelSpecification> AddModelOpticalSpecFromDb(Dictionary<string, ModelSpecification> inputData)
        {
            DataTable sqlTableLot = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = @"Data Source=MSTMS010;Initial Catalog=MES;User Id=mes;Password=mes;";

            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = String.Format(@"SELECT  MODEL_ID,CCT,Vf_min,Vf_max,lm_min,lm_max,lmW_min,CRI_min,CRI_max,[IF] FROM tb_MES_models_ex2;");

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(sqlTableLot);

            foreach (DataRow row in sqlTableLot.Rows)
            {
                string model = row["MODEL_ID"].ToString();

                double vfMin = 0;
                double vfMax = 0;
                double lmMin = 0;
                double lmMax = 0;
                double criMin = 0;
                double lmWMin = 0;
                double cct = 0;
                double criMax = 0;
                double currentF = 0;

                double.TryParse(row["Vf_min"].ToString(), out vfMin);
                double.TryParse(row["Vf_max"].ToString(), out vfMax);
                double.TryParse(row["lm_min"].ToString(), out lmMin);
                double.TryParse(row["lm_max"].ToString(), out lmMax);
                double.TryParse(row["CRI_min"].ToString(), out criMin);
                double.TryParse(row["CRI_max"].ToString(), out criMax);
                double.TryParse(row["lmW_min"].ToString(), out lmWMin);
                double.TryParse(row["CCT"].ToString(), out cct);
                double.TryParse(row["IF"].ToString(), out currentF);


                if (!inputData.ContainsKey(model))
                {
                    inputData.Add(model, new ModelSpecification(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "","","","",null,""));
                }
                double minCct = Math.Round(cct * 0.9, 0);
                double maxCct = Math.Round(cct * 1.1, 0);

                inputData[model].CctMin = minCct;
                inputData[model].CctMax = maxCct;
                inputData[model].Vf_Min = vfMin;
                inputData[model].Vf_Max = vfMax;
                inputData[model].Lm_Min = lmMin;
                inputData[model].Lm_Max = lmMax;
                inputData[model].CRI_Min = criMin;
                inputData[model].CRI_Max = criMax;
                inputData[model].LmW_Min = lmWMin;
                inputData[model].CurrentForward = currentF;
            }

            return inputData;
        }
    }
}
