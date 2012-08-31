using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace grafik2chart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void b1_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string dosya_yolu = openFileDialog1.FileName;
            OleDbConnection con = new OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; " + "data source=" + dosya_yolu + ";Extended Properties=Excel 12.0;");
            con.Open();
            OleDbDataAdapter command = new OleDbDataAdapter("select * from [data$]", con);
            DataTable dt = new DataTable();
            command.Fill(dt);
            int m = dt.Rows.Count;
                 
            double[] y1 = new double[m];
            double[] y2 = new double[m];
            double[] y3 = new double[m];
            double[] y12 = new double[m];
            double[] y26 = new double[m];
            double[] y1226 = new double[m];

            int i = 0;

            
            foreach (DataRow dr in dt.Rows)
            {
                y1[i] = Convert.ToDouble(dt.Rows[i].ItemArray[2].ToString());
                y2[i] = Convert.ToDouble(dt.Rows[i].ItemArray[3].ToString());
                y3[i] = Convert.ToDouble(dt.Rows[i].ItemArray[4].ToString());
                i++;
                //if (i >= 100) break;

            }
            ;
            chart1.Series["Series1"].Points.DataBindY(y1);
            chart1.Series["Series2"].Points.DataBindY(y2);
            for (int k = 0; k < m; k++)
            {
                if (k < 26)
                {
                    y12[k] = 0;
                    y26[k] = 0;
                }
                else
                {
                    for (int l=12 ;  l >= 0; l--)
                    {
                        y12[k] += y1[k - l];
                    }
                    y12[k] /= 12;
                    for (int j=26; j >= 0; j--)
                    {
                        y26[k] += y1[k - j];
                    }
                    y26[k] /= 26;
                }
                y1226[k] = y12[k] - y26[k];
            }
            chart2.Series["Series1"].Points.DataBindY(y1226);
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
