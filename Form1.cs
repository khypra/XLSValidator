using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace XLSValidator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
           InitializeComponent();
        }

        private void ObterXLS(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"C:\",
                    Title = "Abrir arquivo XLS",

                    CheckFileExists = true,
                    CheckPathExists = true,

                    DefaultExt = "xls",
                    Filter = "Arquivos Excel (*.xls)|*.xls",
                    FilterIndex = 1,
                    RestoreDirectory = true,

                    ReadOnlyChecked = true,
                    ShowReadOnly = true
                };

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    var book = new LinqToExcel.ExcelQueryFactory(@"" + openFileDialog1.FileName);
                    if(book != null)
                    {
                        MessageBox.Show("hello");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void LerArquivos()
        {
           
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {}
    }


}
