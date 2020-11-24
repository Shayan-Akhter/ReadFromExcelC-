using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadFromExcelReadSalesQuotation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTableCollection dataTableCollection;
        private void readFileButton_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = ofd.FileName;
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            dataTableCollection = result.Tables;
                            comboBox1.Items.Clear();
                            comboBox3.DataSource = null;
                            comboBox1.Items.AddRange(dataTableCollection.Cast<DataTable>().Select(t => t.TableName).ToArray<string>());
                        }
                    }
                }
            }
        }

        DataTable dt;
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //select data by sheet name
            dt = dataTableCollection[comboBox1.SelectedItem.ToString()];
            var columnNames = (from c in dt.Columns.Cast<DataColumn>()
                               select c.ColumnName).ToArray();
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(columnNames);

            DataTable dtExcel = new DataTable();
            dtExcel = dataTableCollection[comboBox1.SelectedItem.ToString()];
            dataGridView1.Visible = true;
            dataGridView1.DataSource = dtExcel;
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //select data by column name
            if (dt != null)
            {
                string columnName = comboBox2.SelectedItem.ToString();
                var data = dt.DefaultView.ToTable(false, columnName);
                comboBox3.DataSource = data;
                comboBox3.DisplayMember = columnName;
                comboBox3.ValueMember = columnName;
            }
        }
    }
}
