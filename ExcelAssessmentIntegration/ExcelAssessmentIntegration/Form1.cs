using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAssessmentIntegration
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void readExcelSheet(int filterType, string filterChoice)
        {
            //open excelApp and create the new application
            Excel.Application excelApp;
            excelApp = new Excel.Application();
            //workbook 
            Excel.Workbook excelWorkbook;
            //worksheet 
            Excel.Worksheet excelWorksheet;
            //range variable
            Excel.Range range;

            string str;
            int rowCount = 0;
            int colCount = 0;

            //make excel visible to the user
            excelApp.Visible = true;

            string workbookPath = ""; //PATH NEEDS TO BE ADDED/MANIPULATED
            try
            {
                excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                   0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                   true, false, 0, true, false, false);

                //holds the worksheet
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                //read from this sheet
                string currentSheet = "Sheet1";
                excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                //find all valid cells
                range = excelWorksheet.UsedRange;

                //Number of rows
                Console.WriteLine("Number of Rows: " + range.Rows.Count);

                //Number of columns
                Console.WriteLine("Rumber of Columns: " + range.Columns.Count);

                for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                {
                    for (colCount = 1; colCount <= range.Columns.Count; colCount++)
                    {
                        if (range.Cells[rowCount, colCount].Value2 != null)
                        {
                            str = range.Cells[rowCount, colCount].Value2.ToString();
                            Console.WriteLine("Value in cell " + rowCount + " " + colCount + " is " + str); //print each individual cell
                        }

                    }
                }

                //close
                excelWorkbook.Close(true, null, null);
                excelApp.Quit();


            }
            catch (Exception)
            {
                MessageBox.Show("ERROR: FILE NOT READ");
            }



        }

        private void ReadExcelBtn_Click(object sender, EventArgs e)
        {
            int filterType = 0;
            int i = 0;

            System.IO.DirectoryInfo dataFilesDir = new System.IO.DirectoryInfo("..\\dataSheets\\"); //CHANGE DIRECTORY LOCATION???????????
            int filesCount = dataFilesDir.GetFiles().Length;
            MessageBox.Show(filesCount.ToString());

            if (yearRB.Checked)
            {
                filterType = 2;
            }
            if (semesterRB.Checked)
            {
                filterType = 3;
            }
            if (courseRB.Checked)
            {
                filterType = 4;
            }
            if (sectionRB.Checked)
            {
                filterType = 5;
            }

            switch (filterType)
            {
                case 1: //filter by nothing
                    break;
                case 2: //filter by year
                    for (i = 0; i < filesCount; i++)
                    {

                    }
                    break;
            }

        }

       
    }
}

