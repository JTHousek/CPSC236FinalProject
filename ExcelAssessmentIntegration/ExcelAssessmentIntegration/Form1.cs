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
            readExcelSheet(1, "2019");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void filterCriteriaGrpBx_Enter(object sender, EventArgs e)
        {

        }

        private void yearRB_CheckedChanged(object sender, EventArgs e)
        {
            int filterType = 1;
            if (yearRB.Checked == true)
            {
                MessageBox.Show("WorksYears");
            }

            showFilters();
            //filterData(filterType);
        }

        private void semesterRB_CheckedChanged(object sender, EventArgs e)
        {
            int filterType = 2;

            if (semesterRB.Checked == true)
            {
                MessageBox.Show("WorksSem");
            }

            showFilters();
            //filterData(filterType);
        }

        private void courseRB_CheckedChanged(object sender, EventArgs e)
        {
            int filterType = 3;

            if (courseRB.Checked == true)
            {
                MessageBox.Show("WorksCour");
            }

            showFilters();
            //filterData(filterType);
        }

        private void sectionRB_CheckedChanged(object sender, EventArgs e)
        {
            int filterType = 4;

            if (sectionRB.Checked == true)
            {
                MessageBox.Show("WorksSecti");
            }

            showFilters();
           // filterData(filterType);
        }

        public void filterData(int filterCriteria)
        {
            const int LOWESTYEAR = 1950;
            const int HIGHESTYEAR = 3000;

            String getValue;
            int convertVal;


            getValue = filterBoxBx.Text.Trim();

            if (getValue == "")
            {
                MessageBox.Show("Please enter a value in the box");
                filterBoxBx.Clear();
            }

            switch (filterCriteria)
            {
                case 1:
                    if (int.TryParse(getValue, out convertVal))
                    {
                        if (convertVal > LOWESTYEAR || convertVal < HIGHESTYEAR)
                        {
                            //Do filtering here
                        }
                        else
                        {
                            MessageBox.Show("Please enter a valid year to search between 1950 and 3000");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Value is not a number. Please enter a year into the box");
                        filterBoxBx.Clear();
                    }

                    break;
                case 2:
                    break;
                case 3:
                    break;
                case 4:
                    break;
                default:
                    break;
            }
        }

        public void showFilters()
        {
            filterLB.Visible = true;
            filterBoxBx.Visible = true;
            filterBtn.Visible = true;
        }
    }
}

