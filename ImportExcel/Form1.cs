using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Z.Dapper.Plus;

namespace ImportExcel {
    public partial class Form1:Form {
        public Form1() {
            InitializeComponent();
        }

        private void CboSheet_SelectedIndexChanged(object sender,EventArgs e) {

            DataTable dt = tableCollection[CboSheet.SelectedItem.ToString()];

            if(dt != null)
            {

                List<Customer> empresa = new List<Customer>();
                for(int i = 0;i < dt.Rows.Count;i++)
                {
                    Customer customer = new Customer
                    {
                        Tipreg = dt.Rows[i]["Tipreg"].ToString(),
                        Codbdi = dt.Rows[i]["Codbdi"].ToString(),
                        Codneg = dt.Rows[i]["Codneg"].ToString(),
                        Tpmerc = dt.Rows[i]["Tpmerc"].ToString(),
                        Nomres = dt.Rows[i]["Nomres"].ToString(),
                        Especi = dt.Rows[i]["Especi"].ToString(),
                        Prazot = dt.Rows[i]["Prazot"].ToString(),
                        Modref = dt.Rows[i]["Modref"].ToString(),
                        Preabe = decimal.Parse(dt.Rows[i]["Preabe"].ToString()),
                        Premax = decimal.Parse(dt.Rows[i]["Premax"].ToString()),
                        Premin = decimal.Parse(dt.Rows[i]["Premin"].ToString()),
                        Premed = decimal.Parse(dt.Rows[i]["Premed"].ToString()),
                        Preult = decimal.Parse(dt.Rows[i]["Preult"].ToString()),
                        Preofc = decimal.Parse(dt.Rows[i]["Preofc"].ToString()),
                        Preofv = decimal.Parse(dt.Rows[i]["Preofv"].ToString()),
                        Totneg = dt.Rows[i]["Totneg"].ToString(),
                        Quatot = dt.Rows[i]["Quatot"].ToString(),
                        Voltot = decimal.Parse(dt.Rows[i]["Voltot"].ToString()),
                        Preexe = decimal.Parse(dt.Rows[i]["Preexe"].ToString()),
                        Indopc = dt.Rows[i]["Indopc"].ToString(),
                        Fatcot = dt.Rows[i]["Fatcot"].ToString(),
                        Ptoexe = decimal.Parse(dt.Rows[i]["Ptoexe"].ToString()),
                        Codisi = dt.Rows[i]["Codisi"].ToString(),
                        Dismes = decimal.Parse(dt.Rows[i]["Dismes"].ToString())
                    };

                    if (DateTime.TryParseExact(dt.Rows[i]["Datven"].ToString(), "yyyyMMdd", null, DateTimeStyles.AssumeUniversal, out var datven) ||
                        DateTime.TryParseExact(dt.Rows[i]["Datven"].ToString(), "dd/MM/yyyy hh.mm.ss", null, DateTimeStyles.None, out datven))
                        {
                            customer.Datven = datven;
                        }
                    
                    if (DateTime.TryParseExact(dt.Rows[i]["Datpreg"].ToString(), "yyyyMMdd", null, DateTimeStyles.AssumeUniversal, out var datpreg) ||
                        DateTime.TryParseExact(dt.Rows[i]["Datpreg"].ToString(), "dd/MM/yyyy hh.mm.ss", null, DateTimeStyles.None, out datpreg))
                        {
                            customer.Datpreg = datpreg;
                        }

                    empresa.Add(customer);
                }
                empresaBindingSource8.DataSource = empresa;
            }
        }

        DataTableCollection tableCollection;

        private void btnBrowse_Click(object sender,EventArgs e) {
            using(OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97 - 2003 Workbook | *.xls| Excel Workbook|*.xlsx" })
            {
                if(openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    using(var stream = File.Open(openFileDialog.FileName,FileMode.Open,FileAccess.Read))
                    {
                        using(IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }

                            });
                            tableCollection = result.Tables;
                            CboSheet.Items.Clear();
                            foreach(DataTable table in tableCollection)
                                CboSheet.Items.Add(table.TableName);
                        }

                    }
                }
            }
        }

        private void btnImport_Click(object sender,EventArgs e) {
            try
            {
                DapperPlusManager.Entity<Customer>().Table("empresa").BatchSize(1000);

                if(empresaBindingSource8.DataSource is List<Customer> empresa)
                {
                    using(IDbConnection db = new SqlConnection("Data Source= DANIELHOUSE; Database=bdacoes; Integrated Security=True"))
                    {
                        db.UseBulkOptions(options => options.BatchSize = 1000);
                        db.BulkInsert(empresa);
                    }
                    MessageBox.Show("Concluido!!!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"Message",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender,EventArgs e) {

            // TODO: esta linha de código carrega dados na tabela 'bdacoesDataSet7.empresa'. Você pode movê-la ou removê-la conforme necessário.
            //this.empresaTableAdapter6.Fill(this.bdacoesDataSet7.empresa);
            // TODO: esta linha de código carrega dados na tabela 'bdacoesDataSet6.empresa'. Você pode movê-la ou removê-la conforme necessário.
           // this.empresaTableAdapter5.Fill(this.bdacoesDataSet6.empresa);
            // TODO: esta linha de código carrega dados na tabela 'bdacoesDataSet5.empresa'. Você pode movê-la ou removê-la conforme necessário.
           // this.empresaTableAdapter4.Fill(this.bdacoesDataSet5.empresa);
        }

        private void dataGridView1_CellContentClick(object sender,DataGridViewCellEventArgs e) {
        }
    }
}
