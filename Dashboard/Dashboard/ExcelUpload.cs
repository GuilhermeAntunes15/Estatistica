using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel;
using MySql.Data.MySqlClient;

namespace Dashboard
{
    public partial class ExcelUpload : Form
    {

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn
         (
               int nLeftRect,
               int nTopRect,
               int nRightRect,
               int nBottomRect,
               int nWidthEllipse,
               int nHeightEllipse

         );

        private OpenFileDialog openFileDialog1 = new OpenFileDialog
        {
            Title = "Browse Text Files",

            CheckFileExists = true,
            CheckPathExists = true,

            Filter = "Excel Files|*.xls;*.xlsx;*.csv"
        };

        public ExcelUpload()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            pnlNav.Height = btnExcel.Height;
            pnlNav.Top = btnExcel.Top;
            pnlNav.Left = btnExcel.Left;
            btnExcel.BackColor = Color.FromArgb(46, 51, 73);

            txtArquivo.Text = "No files selected";
            btnUpload.Enabled = false;
        }


        private void btnDashbord_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnDashbord.Height;
            pnlNav.Top = btnDashbord.Top;
            pnlNav.Left = btnDashbord.Left;
            btnDashbord.BackColor = Color.FromArgb(46, 51, 73);

            Form1 form1 = new Form1();
            this.Hide();
            form1.Show();
        }

        private void btnAnalytics_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnAnalytics.Height;
            pnlNav.Top = btnAnalytics.Top;
            btnAnalytics.BackColor = Color.FromArgb(46, 51, 73);

        }

        private void btnCalender_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnExcel.Height;
            pnlNav.Top = btnExcel.Top;
            btnExcel.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnContactUs_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnContactUs.Height;
            pnlNav.Top = btnContactUs.Top;
            btnContactUs.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnsettings_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnsettings.Height;
            pnlNav.Top = btnsettings.Top;
            btnsettings.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnDashbord_Leave(object sender, EventArgs e)
        {
            btnDashbord.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnAnalytics_Leave(object sender, EventArgs e)
        {
            btnAnalytics.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnCalender_Leave(object sender, EventArgs e)
        {
            btnExcel.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnContactUs_Leave(object sender, EventArgs e)
        {
            btnContactUs.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnsettings_Leave(object sender, EventArgs e)
        {
            btnsettings.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {
            
            //openFileDialog1.Title = "Select Tasks";
            //openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtArquivo.Text = openFileDialog1.SafeFileName;
                btnUpload.Enabled = true;
                btnSalvar.Enabled = true;
            }
        }

        DataSet ds;

        private void btnUpload_Click(object sender, EventArgs e)
        {
            dgvDados.ColumnCount = 0;
            FileStream fileStream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fileStream);
            reader.IsFirstRowAsColumnNames = true;
            ds = reader.AsDataSet();
            dgvDados.DataSource = ds.Tables[0];
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            string constring = DbConfig.ConnectionString();

            MySqlConnection con = new MySqlConnection(constring);

            con.Open();




            int tamanhoLinhas = dgvDados.Rows.Count;
            foreach (DataGridViewRow row in dgvDados.Rows)
            {
                foreach (DataGridViewColumn col in dgvDados.Columns)
                {
                    if(row.Index == tamanhoLinhas-1)
                    {
                        continue;
                    }
                    // verificar o uso de aspas na hora de inserir 
                    string val = dgvDados.Rows[row.Index].Cells[col.Index].Value.ToString();
                    string insertQuery = "INSERT INTO colunas(valor) VALUES('" + val + "');";

                    MySqlCommand command = new MySqlCommand(insertQuery, con);
                    command.ExecuteNonQuery();
                }
            }
            MessageBox.Show("Data Inserted");
                
          

            con.Close();
        }

        private void ExcelUpload_Load(object sender, EventArgs e)
        {
            dgvDados.ColumnCount = 3;
            dgvDados.Columns[0].Name = "Product ID";
            dgvDados.Columns[1].Name = "Product Name";
            dgvDados.Columns[2].Name = "Product Price";
        }
    }
}
