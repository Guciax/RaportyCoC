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
    public partial class chooseDateOfShipment : Form
    {
        private readonly List<string> listOfShipmentDates;
        public List<string> choosenDates = new List<string>();

        public chooseDateOfShipment(List<string> listOfShipmentDates)
        {
            InitializeComponent();
            this.listOfShipmentDates = listOfShipmentDates;
        }

        private void chooseDateOfShipment_Load(object sender, EventArgs e)
        {
            foreach (var item in listOfShipmentDates)
            {
                checkedListBox1.Items.Add(item, true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (var item in checkedListBox1.CheckedItems)
            {
                choosenDates.Add(item.ToString());
            }
        }
    }
}
