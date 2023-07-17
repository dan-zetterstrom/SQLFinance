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
    public partial class formLogin : Form
    {
        public formMain MainForm;
        public formLogin(formMain mainForm)
        {
            MainForm = mainForm;
            InitializeComponent();
        }

        private void cmdLogin_Click(object sender, EventArgs e)
        {
            if (txtUsername.Text != "")
            {
                if (txtPassword.Text != "")
                {
                    MainForm.UN = txtUsername.Text;
                    MainForm.PW = txtPassword.Text;
                    this.Close();
                }
                else
                {
                    DialogResult result = MessageBox.Show("Are you sure you want to use a blank password?", "Are you sure?", MessageBoxButtons.OKCancel);
                    if (result == DialogResult.OK)
                    {
                        MainForm.UN = txtUsername.Text;
                        MainForm.PW = txtPassword.Text;
                        this.Close();
                    }
                }
            }
            else 
            {
                MessageBox.Show("Username cannot be blank!");
            }
        }

        private void onClose(Object sender, EventArgs e)
        {
            /*
             *Overrides existing onClose method to make sure that
             *the SQL connection is closed when the application closes
             *
             *SQL connection may not close if application is closed in 
             *unorthodox ways
             */
            MainForm.reallyClose = true;
        }
    }
}
