using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exporter
{
    public partial class Form2 : Form
    {
        public string MyProperty { get; private set; }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
        private void login()
        {

            string contraseña = contraseñaTXT.Text;
            if (contraseña == "eliza52" || contraseña == "2424" || contraseña == "9999" || contraseña == "mathias24")
            {
                Form1 frm1 = new Form1();
                frm1.MyProperty = contraseña;
                frm1.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Clave invalida, por favor verifique", "Advertencia");

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            login();
        }

        private void contraseñaTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void contraseñaTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            

            if (e.KeyChar == (char)Keys.Return)
            {
                login();
            }
            






        }

        private void contraseñaTXT_KeyDown(object sender, KeyEventArgs e)
        {
           
        }
    }
}
