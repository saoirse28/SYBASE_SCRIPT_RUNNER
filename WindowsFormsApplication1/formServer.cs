using System;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class formServer : Form
    {
        public string serverName, ipAddress;
        public formServer()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            serverName = textBox1.Text.Trim();
            ipAddress = textBox2.Text.Trim();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void formServer_Load(object sender, EventArgs e)
        {

        }
    }
}
