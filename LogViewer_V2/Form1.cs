﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
            btn_exportToExcel.Enabled = false;
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
            

            if (FileName.ToUpper().Contains("EXECUTION"))
            {
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

                if (dt.Rows.Count > 0)
                    btn_exportToExcel.Enabled = true;
            }
            else
                MessageBox.Show("Unknown File. Please choose Execution log file");
        }

        public void SampleCheckChangedHandler(object objSender, EventArgs ea)
        {
            CheckBox cb = objSender as CheckBox; 

            if (cb.Checked)
                dataGridView1.Columns[cb.Text].Visible = true;
            else
                dataGridView1.Columns[cb.Text].Visible = false;
                
        }

        private void Btn_exportToExcel_Click(object sender, EventArgs e)
        {
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.Filter = "Excel Files | *.xlsx";

            saveFileDialog1.ShowDialog();
            String fileName = saveFileDialog1.FileName;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            // Change properties of the Workbook 

            ExcelApp.Columns.ColumnWidth = 20;

            // Storing header part in Excel
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                ExcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            // Storing Each row and column value to excel sheet
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            
            ExcelApp.Visible = true;
            ExcelApp.ActiveWorkbook.SaveCopyAs(fileName);
            ExcelApp.ActiveWorkbook.Saved = true;
            Marshal.FinalReleaseComObject(ExcelApp);
        }

      
    }
}

