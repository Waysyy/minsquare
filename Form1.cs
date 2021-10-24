using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using Exc = Microsoft.Office.Interop.Excel;
using System.Net;
using System.IO;
using Newtonsoft.Json;

namespace minsquare
{
    
        public partial class Form1 : Form
        {

        
        public Form1()
        {
            InitializeComponent();
            GraphPane pane = zedGraphControl1.GraphPane;
            pane.Title.Text = "График";

        }

        public void Excel()
        {
            try
            {
                string str;
                int rCnt;
                int cCnt;

                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Файл Excel|*.XLSX;*.XLS";
                opf.ShowDialog();
                System.Data.DataTable tb = new System.Data.DataTable();
                string filename = opf.FileName;

                Exc.Application ExcelApp = new Exc.Application();
                Exc._Workbook ExcelWorkBook;
                Exc.Worksheet ExcelWorkSheet;
                Exc.Range ExcelRange;

                ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Exc.XlPlatform.xlWindows, "\t", false,
                    false, 0, true, 1, 0);
                ExcelWorkSheet = (Exc.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                ExcelRange = ExcelWorkSheet.UsedRange;
                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for (cCnt = 1; cCnt <= 2; cCnt++)
                    {
                        str = (string)(ExcelRange.Cells[rCnt, cCnt] as Exc.Range).Text;
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                }
                ExcelWorkBook.Close(true, null, null);
                ExcelApp.Quit();

                releaseObject(ExcelWorkSheet);
                releaseObject(ExcelWorkBook);
                releaseObject(ExcelApp);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Невозможно очистить " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        
        }
        public void Datagrid()
        {
            try
            {
                GraphPane pane = zedGraphControl1.GraphPane;
                pane.CurveList.Clear();
                PointPairList list = new PointPairList();
                PointPairList list1 = new PointPairList();
                PointPairList list2 = new PointPairList();

                string[] x;
                string[] y;
                
                string yd = "";
                string xd = "";
                double n = 0;
                
                double sum1 = 0;
                double sum2 = 0;
                double sum3 = 0;
                double sum4 = 0;

                x = new string[dataGridView1.RowCount-1];
                y = new string[dataGridView1.RowCount-1];
                for (int i = 0; i < dataGridView1.RowCount-1; ++i)
                {
                    x[i] = dataGridView1[0, i].Value.ToString();
                    y[i] = dataGridView1[1, i].Value.ToString();
                    xd = x[i];
                    yd = y[i];
                    sum1 += Convert.ToDouble(xd);
                    sum1 += Convert.ToDouble(yd);
                    sum2 += Convert.ToDouble(xd);
                    sum3 += Convert.ToDouble(yd);
                    sum4 += Convert.ToDouble(xd) * Convert.ToDouble(xd);
                    n = dataGridView1.Rows.Count;
                    list.Add(Convert.ToDouble(xd), Convert.ToDouble(yd));

                    
                }
               

                double a = (sum1 - (sum2 * sum3)) / sum4 - (sum2 * sum2);
                double b = (sum3 - (sum2 * a)) / n;
                double h = (Math.Abs(b - a)) / 100;

                for (double x1 = a; x1 <= b; x1 += h)
                {
                    list1.Add(x1, x1+b);

                }

                LineItem polinom = pane.AddCurve("Sinc", list1, Color.Blue, SymbolType.None);

                for (double x1 = a; x1 <= b; x1 += h)
                {
                    list2.Add(x1, a*(x1*x1)+b*x1);

                }
                LineItem approx = pane.AddCurve("Sinc", list2, Color.Purple, SymbolType.None);

                LineItem dot = pane.AddCurve("Sinc", list, Color.Red, SymbolType.Star);
                dot.Line.IsVisible = false;

                zedGraphControl1.AxisChange();
                zedGraphControl1.Invalidate();

  
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        public void GSheets()
        {
            try
            {
                string link;
                WebRequest req = WebRequest.Create("https://docs.google.com/spreadsheets/d/" + @textBox1.Text + "/gviz/tq");
                WebResponse res = req.GetResponse();
                using (Stream stream = res.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        
                        link = reader.ReadToEnd();

                    }
                }
                link = link.Remove(0, 47);
                var asd = link.LastIndexOf(")");
                link = link.Remove(asd, 2);
                res.Close();

                //var result = JsonConvert.DeserializeObject<List<RootObject>>(link);

                //var result = JsonConvert.DeserializeObject<SortedDictionary<string, C>>(link);

                //C root = JsonConvert.DeserializeObject<C>(link);
                //var result = root;
                //MessageBox.Show(Convert.ToString(result));

                 var result = JsonConvert.DeserializeObject<RootObject>(link);
                dataGridView1.DataSource = result;

                
                
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Datagrid();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            GSheets();
        }
    }
}
