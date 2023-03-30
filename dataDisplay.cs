using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQLFinance
{
    public partial class dataDisplay : Form
    {
        public DataSet DataSet;
        public dataDisplay()
        {
            InitializeComponent();
        }

        public dataDisplay(DataSet dataSet) {
            DataSet = dataSet;
            InitializeComponent();
            try
            {
                dgData.DataSource = DataSet.Tables[0];
                dgData.Update();
            }
            catch (Exception e) {
                MessageBox.Show("No data to display in current date range!");
            }
        }
    }
}
