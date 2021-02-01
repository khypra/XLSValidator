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
using System.Reflection;
using EasyXLS;
using System.Globalization;

namespace AtosCapital
{
    public partial class XLSValidator : Form
    {
        List<Nota> comparativo = new List<Nota>();
        string baseDiretory = "";

        public XLSValidator()
        {
            InitializeComponent();
        }
        static DataTable ConvertToDataTable<T>(List<T> models)
        {
            // creating a data table instance and typed it as our incoming model   
            // as I make it generic, if you want, you can make it the model typed you want.  
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties of that model  
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Loop through all the properties              
            // Adding Column name to our datatable  
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names    
                dataTable.Columns.Add(prop.Name);
            }
            // Adding Row and its value to our dataTable  
            foreach (T item in models)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows    
                    values[i] = Props[i].GetValue(item, null);
                }
                // Finally add value to datatable    
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        //manual da classe externa EasyXLS https://www.easyxls.com/manual/index.html
        public void GenerateExcel(DataTable notas)
        {
            try {
                string diaHoraAgora = DateTime.Now.ToString("dd-MM-yyyy HH-mm");
                ExcelDocument excelFile = new ExcelDocument();
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(notas);
                excelFile.easy_WriteXLSFile_FromDataSet(baseDiretory + "\\Processamento " + diaHoraAgora 
                                                    + ".xls", dataSet, new EasyXLS.ExcelAutoFormat(EasyXLS.Constants.Styles.AUTOFORMAT_EASYXLS1), "Processamento " + diaHoraAgora);
                MessageBox.Show("Processamento Concluido com Sucesso");
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void CompararDados(List<Nota> notas)
        {
            
            foreach(Nota nota in notas)
            {
                using (BancoDados _dbAtos = new BancoDados(true))
                {
                    float soma = 0;
                    var dados = _dbAtos.SqlQueryDataTable(@"select SUM(vlTotalNFe) AS soma from tax.tbNotaFiscal M (NOLOCK) 
                                                            WHERE CONVERT(date,  M.dtEmissao) = '"+ nota.dtNota +"' and M.nrCNPJEmt = '"+ nota.nmArquivo +"'"+
                                                            "GROUP BY M.nrCNPJEmt");
                    foreach(DataRow row in dados.Rows)
                    {
                        soma = float.Parse(row["soma"].ToString());
                    }
                    if (nota.vlContabil == soma || nota.vlMercadoria == soma)
                    {
                        nota.descricao = "ok";
                        nota.vlAtos = soma;
                        comparativo.Add(nota);
                    }
                    else
                    {
                        nota.descricao = "diferença entre valores: Contabil =" + (nota.vlContabil - soma ) + "  :  Mercadoria=" + ( nota.vlMercadoria - soma);
                        nota.vlAtos = soma;
                        comparativo.Add(nota);
                    }
                }
            }

        }

        static string RemoveDiacritics(string text)
        {
            return string.Concat(
                text.Normalize(NormalizationForm.FormD)
                .Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) !=
                                              UnicodeCategory.NonSpacingMark)
              ).Normalize(NormalizationForm.FormC);
        }

        private void ObterXLS(object sender, EventArgs e)
        {
            Filiais temp = new Filiais();
            Filiais aux = new Filiais();
            List<Filiais> filiais = aux.listarFiliais();
            foreach(Filiais f in filiais)
            {
                Console.WriteLine(f.ds_fantasia);
            }
            List<Nota> listaNotas;

            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"C:\",
                    Title = "Abrir arquivo XLS",

                    CheckFileExists = true,
                    CheckPathExists = true,
                    Multiselect = true,

                    DefaultExt = "xls",
                    Filter = "Arquivos Excel (*.xls)|*.xls| Arquivos Excel (*.xlsx)|*.xlsx | Todos (*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,

                    ReadOnlyChecked = true,
                    ShowReadOnly = true
                };

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    foreach (string name in openFileDialog1.FileNames) {
                        listaNotas = new List<Nota>();
                        baseDiretory = Path.GetDirectoryName(name);
                        string nameArqv = Path.GetFileNameWithoutExtension(name);
                        nameArqv = RemoveDiacritics(nameArqv);
                        try{
                            temp = null;
                            temp = filiais.Find(x => x.ds_fantasia.Contains(nameArqv));
                            if(temp == null)
                            {
                                MessageBox.Show("Nome do Arquivo não refere a uma Filial");
                                break;
                            }
                        }
                        catch (Exception exe)
                        {
                            MessageBox.Show(exe.Message);
                        }
                        var file = new LinqToExcel.ExcelQueryFactory(@"" + name);
                        if (file != null)
                        {

                            var querry =
                                from row in file.Worksheet("Sheet")
                                let item = new
                                {
                                    vlMercadoria = row["Valor mercadorias"].Cast<float>(),
                                    vlContabil = row["Valor contábil"].Cast<float>(),
                                    dtNota = row["Data do doc#"].Cast<DateTime>(),
                                }
                                select item;
                            foreach (var g in querry)
                            {
                                listaNotas.Add(new Nota
                                {
                                    nmArquivo = temp.nu_cnpj,
                                    vlMercadoria = g.vlMercadoria,
                                    vlContabil = g.vlContabil,
                                    dtNota = g.dtNota.ToString("yyyyMMdd"),
                                    descricao = "",
                                    vlAtos = 0
                                }) ;

                            }

                            try
                            {
                                listaNotas = Nota.notasAgrupadas(listaNotas);
                                CompararDados(listaNotas);
                            }
                            catch(Exception err)
                            {
                                MessageBox.Show(err.Message);
                            }

                        }
                        
                    }
                    //depois de rodar todas os arquivos converte a lista em DataTable e manda para a conversão em xls
                    DataTable result = ConvertToDataTable<Nota>(comparativo);
                    GenerateExcel(result);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }


}
