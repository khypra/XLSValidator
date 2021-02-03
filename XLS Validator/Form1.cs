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
        static List<Nota> comparativo = new List<Nota>();
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
                string diaHoraAgora = DateTime.Now.ToString("dd-MM-yy HH-mm-ss");
                ExcelDocument excelFile = new ExcelDocument();
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(notas);
                excelFile.easy_WriteXLSFile_FromDataSet(baseDiretory + "\\Processamento " + diaHoraAgora 
                                                    + ".xls", dataSet, new EasyXLS.ExcelAutoFormat(EasyXLS.Constants.Styles.AUTOFORMAT_EASYXLS1), "Processamento"+ diaHoraAgora);
                MessageBox.Show("Processamento Concluido com Sucesso");
                comparativo = new List<Nota>();
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
                label1.Text = "Lendo o CNPJ: " + nota.nmArquivo;
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
                        progressBar1.PerformStep();
                    }
                    else
                    {
                        nota.descricao = "diferença entre valores: Contabil =" + (nota.vlContabil - soma ) + "  :  Mercadoria=" + ( nota.vlMercadoria - soma);
                        nota.vlAtos = soma;
                        comparativo.Add(nota);
                        progressBar1.PerformStep();
                    }
                }

            }

        }
        private void CompararNotas(List<Nota> notas)
        {
            List<Nota> compara = new List<Nota>();
            Nota temp2;
            label2.Text = "Comparando " + notas[0].nmArquivo + " com o banco da Atos";
            using (BancoDados _dbAtos = new BancoDados(true))
            {
                var dados = _dbAtos.SqlQueryDataTable(@"select M.vlTotalNFe, M.nrChaveAcesso AS nrChave from tax.tbNotaFiscal M (NOLOCK) 
                                                            WHERE CONVERT(date,  M.dtEmissao) = '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and M.nrCNPJEmt = '" + notas[0].nmArquivo + "';");
                    
                
                foreach (DataRow row in dados.Rows)
                {
                    compara.Add(new Nota
                    { 
                      nmArquivo = notas[0].nmArquivo,
                      dtNota = dateTimePicker1.Value.ToString("yyyyMMdd"),
                      vlAtos =  float.Parse(row["vltotalNFe"].ToString()), 
                      cdChave = row["nrChave"].ToString() 
                    });
                }

                foreach(Nota n in compara)
                {
                    temp2 = new Nota();
                    temp2 = notas.Find(x => x.cdChave == n.cdChave);
                    if (temp2 != null)
                    {
                        if (temp2.vlContabil.Equals(n.vlAtos))
                        {
                            notas.Remove(temp2);
                            progressBar2.PerformStep();
                        }
                        else
                        {
                            n.vlMercadoria = temp2.vlMercadoria;
                            n.vlContabil = temp2.vlContabil;
                            n.descricao = "diferença entre vlCliente = "+ temp2.vlContabil +" e vlAtos = " + n.vlAtos;
                            comparativo.Add(n);
                            notas.Remove(temp2);
                            progressBar2.PerformStep();
                        }
                    }
                    else
                    {
                        n.descricao = "Nota existe na atos e não existe no cliente";
                        comparativo.Add(n);
                        progressBar2.PerformStep();
                    }
                }
                foreach(Nota n in notas)
                {
                    n.descricao = "Nota existe no cliente e não existe na atos";
                    comparativo.Add(n);
                    progressBar2.PerformStep();
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
                    foreach (string name in openFileDialog1.FileNames) 
                    {
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
                                label1.Text = "Lendo arquivos: ";
                                progressBar1.Maximum = listaNotas.Count();
                                progressBar1.Minimum = 0;
                                progressBar1.Value = 0;
                                progressBar1.Step = 1;
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
                comparativo = new List<Nota>();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CompararDiaErro(object sender, EventArgs e)
        {
            Filiais temp = new Filiais();
            Filiais aux = new Filiais();
            List<Filiais> filiais = aux.listarFiliais();
            foreach (Filiais f in filiais)
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
                    Multiselect = false,

                    DefaultExt = "xls",
                    Filter = "Arquivos Excel (*.xls)|*.xls| Arquivos Excel (*.xlsx)|*.xlsx | Todos (*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,

                    ReadOnlyChecked = true,
                    ShowReadOnly = true
                };

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    foreach (string name in openFileDialog1.FileNames)
                    {
                        listaNotas = new List<Nota>();
                        baseDiretory = Path.GetDirectoryName(name);
                        string nameArqv = Path.GetFileNameWithoutExtension(name);
                        nameArqv = RemoveDiacritics(nameArqv);
                        try
                        {
                            temp = null;
                            temp = filiais.Find(x => x.ds_fantasia.Contains(nameArqv));
                            if (temp == null)
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
                                    cdChave = row["Chave acesso"].Cast<String>(),
                                    vlMercadoria = row["Valor mercadorias"].Cast<float>(),
                                    vlContabil = row["Valor contábil"].Cast<float>(),
                                    dtNota = row["Data do doc#"].Cast<DateTime>(),
                                }
                                select item;
                            foreach (var g in querry)
                            {
                                listaNotas.Add(new Nota
                                {
                                    cdChave = g.cdChave,
                                    nmArquivo = temp.nu_cnpj,
                                    vlMercadoria = g.vlMercadoria,
                                    vlContabil = g.vlContabil,
                                    dtNota = g.dtNota.ToString("yyyyMMdd"),
                                    descricao = "",
                                    vlAtos = 0
                                });

                            }

                            try
                            {
                                listaNotas = Nota.notasAgrupadasPorChave(listaNotas, dateTimePicker1.Value);
                                label2.Text = "Comparando Arquivo: ";
                                progressBar2.Maximum = listaNotas.Count();
                                progressBar2.Minimum = 0;
                                progressBar2.Value = 0;
                                progressBar2.Step = 1;
                                CompararNotas(listaNotas);
                            }
                            catch (Exception err)
                            {
                                MessageBox.Show(err.Message);
                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GerarRelatorio(object sender, EventArgs e)
        {
            DataTable result = ConvertToDataTable<Nota>(comparativo);
            GenerateExcel(result);
        }
    }

   
}

