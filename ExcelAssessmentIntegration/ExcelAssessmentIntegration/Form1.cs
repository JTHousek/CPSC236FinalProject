using System;
using System.IO;
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
    public partial class ExcelIntegrationAssessmentWindow : Form
    {
        System.IO.DirectoryInfo dataFilesDir = new System.IO.DirectoryInfo("..\\dataSheets\\"); //CHANGE DIRECTORY LOCATION???????????
        public ExcelIntegrationAssessmentWindow()
        {
            InitializeComponent();
        }
        private void readExcelSheet(string sheetPath)
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
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            //change
            string workbookPath = path + "\\bin\\dataSheets\\" + sheetPath;
            MessageBox.Show(workbookPath);

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
                            MessageBox.Show("Value in cell " + rowCount + " " + colCount + " is " + str); //print each individual cell
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

            String[] delimitedFileName;
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
                        FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[0] == yearTBx.Text)
                            {
                                readExcelSheet(file.Name);
                            }
                        }
                    }
                    break;
                case 3: //filter by semester
                    for (i = 0; i < filesCount; i++)
                    {
                        FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[1] == semesterCmBx.SelectedItem)
                            {
                                readExcelSheet(file.Name);
                            }
                        }
                    }
                    break;
                case 4: //filter by course
                    for (i = 0; i < filesCount; i++)
                    {
                        FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[2] == courseCmBx.SelectedItem)
                            {
                                readExcelSheet(file.Name);
                            }
                        }
                    }
                    break;
                case 5: //filter by year
                    for (i = 0; i < filesCount; i++)
                    {
                        FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[3] == sectionCmBx.SelectedItem)
                            {
                                readExcelSheet(file.Name);
                            }
                        }
                    }
                    break;
            }

        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void filterCriteriaGrpBx_CheckedChanged(object sender, EventArgs e)
        {
            if (yearRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;
            }
            if (semesterRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;
            }
            if (courseRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionCmBx.Enabled = false;
            }
            if (sectionRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionCmBx.Enabled = true;
            }

        }
    }
}