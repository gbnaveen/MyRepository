using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ProjExcelOps
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Approach 1: Using MS Office componenets.
                //ReadExcelFile();





                //foreach (DataRow dr in dtData.Rows)
                //{
                //    string str1 = dr["Candidate"].ToString();
                //    //Do what you need to do with your data here

                //}

        }

        public void ReadExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Naveen\Projects\ProjExcelOps\ProjExcelOps\Internshipdetails.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= rowCount; i++)
            {
                txtCandName.Text = xlRange.Cells[i, 3].Value2.ToString();
                txtGender.Text = xlRange.Cells[i, 4].Value2.ToString();
                txtQual.Text = xlRange.Cells[i, 5].Value2.ToString();
                txtRegType.Text = xlRange.Cells[i, 6].Value2.ToString();

                //for (int j = 2; j <= colCount; j++)
                //{
                //    //new line
                //    if (j == 1)
                //        Console.Write("\r\n");

                //    //write the value to the Form Controls.
                //    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                //    {
                //        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                //        txtCandName.Text = xlRange.Cells[i, j].Value2.ToString();

                //    }

                //}
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Approach 2: using OleDB
            Class1 objcls1 = new Class1();
            DataTable dtData = objcls1.GetExcel();
            dataGridView1.DataSource = dtData.DefaultView;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                txtCandName.Text   = dataGridView1.SelectedRows[0].Cells["Candidate Name"].Value.ToString();
                txtGender.Text   = dataGridView1.SelectedRows[0].Cells["Gender"].Value.ToString();
                txtQual.Text = dataGridView1.SelectedRows[0].Cells["Highest Qualification"].Value.ToString();
                txtRegType.Text = dataGridView1.SelectedRows[0].Cells["Req Type (FTE or Direct -Contract OR 3rd Party Contract)"].Value.ToString();
            }
        }

        //private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    MessageBox.Show("Hi");
        //}
    }
}
