using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace RaportyCoC
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private List<string> ChartPointTypes()
        {
            List<string> result = new List<string>();
            foreach (var ChartType in Enum.GetValues(typeof(MarkerStyle)))
            {
                result.Add(Convert.ToString(ChartType));
            }
            //Add an option the the dropdown menu
            // Convert.ToString(ChartType) <- Text of Item
            // Convert.ToInt32(ChartType) <- Value of Item
            return result;
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            string signature = System.Configuration.ConfigurationManager.AppSettings["Podpis"];
            string defaultPath = System.Configuration.ConfigurationManager.AppSettings["Folder"];

            textBox1.Text = signature;
            textBox2.Text = defaultPath;
            comboBox1.Items.AddRange(ChartPointTypes().ToArray());
        }

        public static void AddOrUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            AddOrUpdateAppSettings("Podpis", textBox1.Text);
            AddOrUpdateAppSettings("Folder", textBox2.Text);
            this.Close();
        }
    }
}
