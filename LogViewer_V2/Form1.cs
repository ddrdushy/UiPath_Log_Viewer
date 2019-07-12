using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace LogViewer_V2
{
    
    public partial class Form1 : Form
    {
        DataTable dt ;

        public Form1()
        {
            InitializeComponent();
            dt = new DataTable();
            Type type = typeof(logClass);
            var properties = type.GetProperties();

            CheckBox box;
            int count = 0;
            int count2 = 20;
            foreach (var item in properties)
            {

                dt.Columns.Add(new DataColumn(item.Name));

                box = new CheckBox();
                box.Tag = item.Name.ToString();
                box.Text = item.Name;
                box.AutoSize = true;
                //box.Location = new Point(10, count * 10); //vertical
                if (count == 6)
                {
                    count = 0;
                    count2 += 30;
                }

                box.Location = new Point(count * 200, count2); //horizontal
                box.Checked = true;
                box.CheckedChanged += SampleCheckChangedHandler;
                this.gb1.Controls.Add(box);
                count++;
            }


            gb1.AutoSize = true;
            dataGridView1.Location = new Point(5, count2 + 75);
            dataGridView1.ScrollBars = ScrollBars.Both;
            dataGridView1.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Left);
            dataGridView1.DataSource = dt;

            btn_exportToExcel.AutoSize = true;
            btn_exportToExcel.Anchor = (AnchorStyles.Top|AnchorStyles.Right);
            btn_exportToExcel.Size = new Size(btn_submit.Width, gb1.Height - 8);
            

        }

        private void Btn_submit_Click(object sender, EventArgs e)
        {
            
            String FileName = txt_fileName.Text;

            using (FileStream fs = File.Open(FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (BufferedStream bs = new BufferedStream(fs))
                {
                    using (StreamReader sr = new StreamReader(bs))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            logClass log = JsonConvert.DeserializeObject<logClass>(line.Substring(line.IndexOf("{")));
                            log.logDate = line.Split(' ')[0];

                            dt.Rows.Add(log.GetAllPropertyValues().ToArray());
                            dataGridView1.Refresh();
                        }
                    }
                }
            }

        }

        public void SampleCheckChangedHandler(object objSender, EventArgs ea)
        {
            CheckBox cb = objSender as CheckBox; 

            if (cb.Checked)
                dataGridView1.Columns[cb.Text].Visible = true;
            else
                dataGridView1.Columns[cb.Text].Visible = false;
                
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void Btn_exportToExcel_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;


            xlexcel = new Microsoft.Office.Interop.Excel.Application();

            xlexcel.Visible = true;
            xlexcel.DisplayAlerts = false;

            xlWorkBook = xlexcel.Workbooks.Add(Type.Missing);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.ActiveSheet;
            xlWorkSheet.Name = "Test work sheet";

            excelCellrange = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[dt.Rows.Count, dt.Columns.Count]];
            excelCellrange.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;


            //Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            //CR.Select();
            //xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

        }
    }
}
