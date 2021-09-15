using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace XMLtoXLSX
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult drResult = OFD.ShowDialog();
            if (drResult == System.Windows.Forms.DialogResult.OK)
                txtXmlPath.Text = OFD.FileName;
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
                if (chbExcelFileName.Checked && txtFileName.Text != "" && txtXmlPath.Text != "")
                {
                    if (File.Exists(txtXmlPath.Text))
                    {
                        string CustXmlFilePath =
                               Path.Combine(new FileInfo(txtXmlPath.Text).DirectoryName, txtFileName.Text);
                        System.Data.DataTable dt = CreateDataTableFromXml(txtXmlPath.Text);
                        ExportDataTableToExcel(dt, CustXmlFilePath);
                        MessageBox.Show("Dönüştürme başarılı!!");
                    }

                }
                else if (!chbExcelFileName.Checked || txtXmlPath.Text != "")
                {
                    if (File.Exists(txtXmlPath.Text))
                    {
                        FileInfo fi = new FileInfo(txtXmlPath.Text);
                        string XlFile = fi.DirectoryName + "\\" + fi.Name.Replace
                        (fi.Extension, ".xlsx");
                        System.Data.DataTable dt = CreateDataTableFromXml
                        (txtXmlPath.Text);
                        ExportDataTableToExcel(dt, XlFile);
                        MessageBox.Show("Dönüştürme başarılı!!");
                    }
                    else
                    {
                        MessageBox.Show("Lütfen bir XML dosyası seçiniz");
                    }
                }
                else
                {
                    MessageBox.Show("Lütfen gerekli parametreleri doğru giriniz!!");
                }
            }

        private void ExportDataTableToExcel(System.Data.DataTable table, string Xlfile)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook book = excel.Application.Workbooks.Add(Type.Missing);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
            excelWorkSheet.Name = table.TableName;

            progressBar1.Maximum = table.Columns.Count;
            for (int i = 1; i < table.Columns.Count + 1; i++)
            {
                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                if (progressBar1.Value < progressBar1.Maximum)
                {
                    progressBar1.Value++;
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            progressBar1.Maximum = table.Rows.Count;
            for (int j = 0; j < table.Rows.Count; j++)
            {
                for (int k = 0; k < table.Columns.Count; k++)
                {
                    excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                }

                if (progressBar1.Value < progressBar1.Maximum)
                {
                    progressBar1.Value++;
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            book.SaveAs(Xlfile);
            book.Close(true);
            excel.Quit();
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(excel);
        }
        public System.Data.DataTable CreateDataTableFromXml(string XmlFile)
        {
            System.Data.DataTable Dt = new System.Data.DataTable();
            try
            {
                DataSet ds = new DataSet();
                ds.ReadXml(XmlFile);
                Dt.Load(ds.CreateDataReader());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex);
            }
            return Dt;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OFD.Filter = "Xml Dosyaları (.xml)|*.xml";
            this.BackColor =ColorTranslator.FromHtml("#1c273a");
        }

        private void chbExcelFileName_CheckedChanged(object sender, EventArgs e)
        {
            switch (chbExcelFileName.Checked)
            {
                case true:
                    txtFileName.Enabled = true;
                    break;
                case false:
                    txtFileName.Enabled = false;
                    break;
            }
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://cyberyazilim.com");
        }

        private void rdrXML_CheckedChanged(object sender, EventArgs e)
        {
            //labelControl1.Text = "XML Dosyasını Seçiniz";
            //chbExcelFileName.Text = "Excel Dosya Adı";
            //OFD.Filter = "Xml Dosyaları (.xml)|*.xml";
        }

        private void rdrExcel_CheckedChanged(object sender, EventArgs e)
        {
            //labelControl1.Text = "Excel Dosyasını Seçiniz";
            //chbExcelFileName.Text = "XML Dosya Adı";
            //OFD.Filter = "Excel Dosyaları (.xlsx)|*.xlsx";
        }
    }
}
