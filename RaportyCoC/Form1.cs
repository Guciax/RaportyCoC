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
        string currentModel = "";
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Dictionary<string, PcbTesterMeasurements> measurementsDictionary = SqlOpeations.GetMeasurementsForBoxes(rTxtBox.Lines);
            DataTable sdcmTable = SdcmCalculation.makeSdcmTable(measurementsDictionary, modelSPecification, ref currentModel);

            Charting.DrawHistogramChart(chart1, measurementsDictionary.Select(val => val.Value.Lm).ToList(), modelSPecification[currentModel].Lm_Min, modelSPecification[currentModel].Lm_Max);

            dataGridView1.DataSource = sdcmTable;
        }
    }
}
