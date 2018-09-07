using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace RaportyCoC
{
    public partial class Form1 : Form
    {
        Dictionary<string, ModelSpecification> modelSPecification = new Dictionary<string, ModelSpecification>();
        Dictionary<string, PcbTesterMeasurements> measurementsDictionary = new Dictionary<string, PcbTesterMeasurements>();
        List<string> boxes = new List<string>();
        string currentModel = "";
        string defaultPath = "";
        string user = "";

        public Form1()
        {
            InitializeComponent();
        }


        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            modelSPecification = excelOperations.loadExcel();
            modelSPecification = SqlOpeations.AddModelOpticalSpecFromDb(modelSPecification);

            user = System.Configuration.ConfigurationManager.AppSettings["Podpis"];
            defaultPath = System.Configuration.ConfigurationManager.AppSettings["Folder"];
#if DEBUG
            chart1.Visible = true;
#endif 
        }

        private void GetMeasuremntsForBoxOrPallet()
        {
            boxes = new List<string>();

            foreach (var line in rTxtBox.Lines)
            {
                if (line.Trim() != "")
                {
                    boxes.Add(line.Trim());
                }
            }

            if (!radioBox.Checked)
            {
                boxes = SqlOpeations.GetBoxesIdFromPalletId(boxes.ToArray());
            }
            if (boxes.Count > 0)
            {
                measurementsDictionary = SqlOpeations.GetMeasurementsForBoxes(boxes.ToArray());
            }
            else
            {
                MessageBox.Show("Brak tego ID w bazie.");
                button2.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!radioButtonOrder.Checked)
            {
                GetMeasuremntsForBoxOrPallet();
            }
            else
            {

                    List<string> ordersList = new List<string>();
                    foreach (var item in rTxtBox.Lines)
                    {
                        string order = item.Trim();
                        if (order!="")
                        {
                            ordersList.Add(order);
                        }
                    }
                    measurementsDictionary = SqlOpeations.GetMeasurementsForOrderNo(ordersList);

            }

            if (measurementsDictionary.Count > 0)
            {
                if (SdcmCalculation.CheckTestMeasurements(measurementsDictionary, modelSPecification, ref currentModel, dataGridView1, labelInfo))
                {
                    button2.Enabled = true;
                    panel1.BackColor = Color.Lime;
                    
#if DEBUG
                        Charting.DrawHistogramChart(chart1, measurementsDictionary.Select(val => val.Value.Vf).ToList(), modelSPecification[currentModel].Vf_Min, modelSPecification[currentModel].Vf_Max);
                        
#endif
                    if (!IfThisModelIsTridonic(currentModel))
                    {
                        MessageBox.Show("To nie jest model Tridoic!");
                        button2.Enabled = false;
                    }
                }
                else
                {
                    button2.Enabled = false;
                }

                //Charting.DrawHistogramChart(chart1, measurementsDictionary.Select(val => val.Value.Lm).ToList(), modelSPecification[currentModel].Lm_Min, modelSPecification[currentModel].Lm_Max);

            }
            else
            {
                MessageBox.Show("Brak tego ID w bazie.");
                button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excelOperations.CreateExcelReport(measurementsDictionary, modelSPecification[currentModel], boxes.ToArray(), user, defaultPath, dateTimePicker1.Value);
        }

        private static bool IfThisModelIsTridonic(string model)
        {
            if (model.Contains("-"))
            {
                string[] familyList = new string[] { "D1", "E2", "K1", "K2", "G1", "G2", "K5", "G5", "22", "33", "53" };
                string family = model.Split('-')[0].Replace("LLFML","");
                if (familyList.Contains(family))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Settings settingsForm = new Settings();
            if(settingsForm.ShowDialog()== DialogResult.OK)
            {
                user = System.Configuration.ConfigurationManager.AppSettings["Podpis"];
                defaultPath = System.Configuration.ConfigurationManager.AppSettings["Folder"];
            }
        }

        private void rTxtBox_TextChanged(object sender, EventArgs e)
        {
            panel1.BackColor = Color.LightSteelBlue;
        }
    }
}
