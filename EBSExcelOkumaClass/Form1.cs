using System;
using System.Windows.Forms;

namespace EBSExcelOkumaClass
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OpenFileDialog op;
        ExcellRead er = new ExcellRead();
        private void button1_Click(object sender, EventArgs e)
        {
            er.Excelverioku("select * from [" + comboBox1.Text + "$]", dataGridView1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            op = new OpenFileDialog();
            if (op.ShowDialog() == DialogResult.OK)
            {
                ExcellRead.yol = op.FileName.ToString();
                ExcellRead.GetEBSSayfaAdiAl(comboBox1);

            }

        }
    }
}
