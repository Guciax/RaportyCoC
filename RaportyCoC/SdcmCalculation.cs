using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RaportyCoC
{
    class SdcmCalculation
    {
        private static double CalculateSDCM(PcbTesterMeasurements testResults, ModelSpecification modelSPecification)
        {
            double result = 0;
            if (modelSPecification.A > 0)
            {
                double partA = Math.Pow((modelSPecification.Cx - testResults.Cx) * modelSPecification.A - (modelSPecification.Cy - testResults.Cy) * modelSPecification.B, 2) / Math.Pow(modelSPecification.C, 2);
                double partB = Math.Pow((modelSPecification.Cx - testResults.Cx) * modelSPecification.B + (modelSPecification.Cy - testResults.Cy) * modelSPecification.A, 2) / Math.Pow(modelSPecification.D, 2);
                double AllensFactor = 0.04;
                result = Math.Sqrt(partA + partB);
                if (result > 0.04)
                {
                    result = result - AllensFactor;
                }
            }
            return result;
        }

        public static bool CheckTestMeasurements(Dictionary<string, PcbTesterMeasurements> testResults, Dictionary<string, ModelSpecification> modelSPecification, ref string currentModel, DataGridView grid, Label infoLabel)
        {
            bool result = true;
            DataTable ngResult = new DataTable();
            DataTable okResult = new DataTable();

            HashSet<string> modelCheck = new HashSet<string>();
            ngResult.Columns.Add("serial_No");
            Dictionary<string, int[]> indexOfNg = new Dictionary<string, int[]>();

            ngResult.Columns.Add("SDCM", typeof(double));
            ngResult.Columns.Add("Vf", typeof(double));
            ngResult.Columns.Add("lm", typeof(double));
            ngResult.Columns.Add("lm_w", typeof(double));
            ngResult.Columns.Add("CRI", typeof(double));
            ngResult.Columns.Add("CCT", typeof(double));
            ngResult.Columns.Add("Cx", typeof(double));
            ngResult.Columns.Add("Cy", typeof(double));
            ngResult.Columns.Add("WYNIK");
            okResult = ngResult.Clone();

            foreach (var testedPcb in testResults)
            {
                string model = testedPcb.Value.Model;
                modelCheck.Add(model);
                double sdcm = 0;
                ModelSpecification modelSpec = null;
                int[] ngIndex = new int[] { 0, 0, 0, 0, 0, 0 };

                if (modelSPecification.TryGetValue(model, out modelSpec))
                {
                    bool allOK = true;
                    sdcm = CalculateSDCM(testedPcb.Value, modelSpec);

                    if (sdcm > modelSPecification[model].MaxSdcm)
                    {
                        allOK = false;
                    }
                    if (testedPcb.Value.Vf < modelSPecification[model].Vf_Min || testedPcb.Value.Vf > modelSPecification[model].Vf_Max)
                    {
                        allOK = false;

                    }
                    if (testedPcb.Value.Lm < modelSPecification[model].Lm_Min || testedPcb.Value.Lm > modelSPecification[model].Lm_Max)
                    {
                        allOK = false;
                    }
                    if (testedPcb.Value.Cct < modelSPecification[model].CctMin || testedPcb.Value.Cct > modelSPecification[model].CctMax)
                    {
                        allOK = false;
                    }
                    if (testedPcb.Value.LmW < modelSPecification[model].LmW_Min)
                    {
                        allOK = false;
                    }
                    if (testedPcb.Value.Cri < modelSPecification[model].CRI_Min)
                    {
                        allOK = false;
                    }

                    string testResult = "OK";
                    if (!allOK)
                    {
                        testResult = "NG";
                        ngResult.Rows.Add(testedPcb.Key, sdcm, testedPcb.Value.Vf, testedPcb.Value.Lm, testedPcb.Value.LmW, testedPcb.Value.Cri, testedPcb.Value.Cct, testedPcb.Value.Cx, testedPcb.Value.Cy, testResult);
                        result = false;
                    }
                    else
                    {
                        okResult.Rows.Add(testedPcb.Key, sdcm, testedPcb.Value.Vf, testedPcb.Value.Lm, testedPcb.Value.LmW, testedPcb.Value.Cri, testedPcb.Value.Cct, testedPcb.Value.Cx, testedPcb.Value.Cy, testResult);
                    }

                    
                    //result.Rows.Add(testedPcb.Key, model, modelSPecification[model].Cx + "x" + modelSPecification[model].Cy + "  CCT="+ modelSPecification[model].Cct+"K" , testedPcb.Value.Cx, testedPcb.Value.Cy, modelSPecification[model].MaxSdcm, sdcm);
                    
                }
                else
                {
                    if (File.Exists(@"Y:\Manufacturing_Center\Integral Quality Management\Dane CofC.xlsx"))
                    {
                        MessageBox.Show("Brak modelu: " + model + @" w pliku Excel: Y:\Manufacturing_Center\Integral Quality Management\Dane CofC.xlsx");
                    }
                    else
                    {
                        MessageBox.Show(@"Brak dostępu do pliku Y:\Manufacturing_Center\Integral Quality Management\Dane CofC.xlsx" + Environment.NewLine + "Sprawdź czy podłączone są dyski sieciowe");
                    }
                    break;
                }

                

            }

            if (modelCheck.Count > 1)
            {
                string msg = "Uwaga wykryto pomeszane modele!" + Environment.NewLine;
                foreach (var mdl in modelCheck)
                {
                    msg += mdl + Environment.NewLine;
                    currentModel = "Kilka modeli!!";
                    result = false;
                }
                MessageBox.Show(msg);
            }
            else if (modelCheck.Count > 0)
            {
                currentModel = modelCheck.ToList()[0];
            }

            //result.DefaultView.Sort = "Wynik_SDCM DESC";
            grid.DataSource = ngResult;
            infoLabel.Text = "Model: " + currentModel + Environment.NewLine + "OK: " + okResult.Rows.Count + Environment.NewLine + "NG: " + ngResult.Rows.Count;

            return result;
        }
    }
}
