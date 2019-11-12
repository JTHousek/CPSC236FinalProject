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
        System.IO.DirectoryInfo dataFilesDir = new System.IO.DirectoryInfo("..\\dataSheets\\");
        private Node obj1 = new Node();
        private Node obj2 = new Node();
        private Node obj3 = new Node();
        public ExcelIntegrationAssessmentWindow()
        {
            InitializeComponent();
        }
        private void readExcelSheet(string sheetPath)
        {
            String objStr = "";
            int convVal = 0;
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
            //MessageBox.Show(workbookPath);

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
                        objStr = range.Cells[rowCount, 1].Value2.ToString();
                        str = range.Cells[rowCount, colCount].Value2.ToString();

                        switch(objStr)
                        {
                            case "objective1":
                                if (range.Cells[rowCount, colCount].Value2 != null)
                                {
                                    str = range.Cells[rowCount, colCount].Value2.ToString();
                                    if (int.TryParse(str, out convVal))
                                    {
                                        switch(colCount)
                                        {
                                            case 2:
                                                obj1.setStudents(obj1.getStudents() + convVal);
                                                break;

                                            case 3:
                                                obj1.setStudents(obj1.getMaxScore() + convVal);
                                                break;

                                            case 4:
                                                obj1.setStudents(obj1.getActualScore() + convVal);
                                                break;

                                            default:
                                                MessageBox.Show("no");
                                                break;
                                        }
                                    }
                                    
                                }
                                break;
                            case "objective2":
                                if (range.Cells[rowCount, colCount].Value2 != null)
                                {
                                    str = range.Cells[rowCount, colCount].Value2.ToString();
                                    if (int.TryParse(str, out convVal))
                                    {
                                        switch (colCount)
                                        {
                                            case 2:
                                                obj2.setStudents(obj2.getStudents() + convVal);
                                                break;

                                            case 3:
                                                obj2.setStudents(obj2.getMaxScore() + convVal);
                                                break;

                                            case 4:
                                                obj2.setStudents(obj2.getActualScore() + convVal);
                                                break;

                                            default:
                                                MessageBox.Show("no");
                                                break;
                                        }
                                    }

                                }
                                break;
                            case "objective3":
                                if (range.Cells[rowCount, colCount].Value2 != null)
                                {
                                    str = range.Cells[rowCount, colCount].Value2.ToString();
                                    if (int.TryParse(str, out convVal))
                                    {
                                        switch (colCount)
                                        {
                                            case 2:
                                                obj3.setStudents(obj3.getStudents() + convVal);
                                                break;

                                            case 3:
                                                obj3.setStudents(obj3.getMaxScore() + convVal);
                                                break;

                                            case 4:
                                                obj3.setStudents(obj3.getActualScore() + convVal);
                                                break;

                                            default:
                                                MessageBox.Show("no");
                                                break;
                                        }
                                    }

                                }
                                break;
                        }

                    }
                }

                MessageBox.Show(obj1.getStudents().ToString());

                //close
                excelWorkbook.Close(true, null, null);
                excelApp.Quit();


            }
            catch (FileNotFoundException ex)
            {
                consoleOutputTxB.AppendText("ERROR: FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }
        }

        private void ReadExcelBtn_Click(object sender, EventArgs e)
        {
            int filterType = 0;
            int convertYear = 0;
            int convertSection = 0;
            string semesterSelected;
            const int FINAL_YEAR = 99;
            const int ZERO = 00;

            String[] delimitedFileName;
            int filesCount = dataFilesDir.GetFiles().Length;
            //MessageBox.Show(filesCount.ToString());

            if (noneRB.Checked)
            {
                filterType = 1;
            }
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

            FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files

            switch (filterType)
            {
                case 1: //filter by nothing    
                    
                    foreach (FileInfo file in Files)
                    {
                        readExcelSheet(file.Name);  
                    }
                    
                    break;

                case 2: //filter by year 

                    if (Int32.TryParse(yearTBx.Text.Trim(), out convertYear))
                    {
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (convertYear >= ZERO && convertYear <= FINAL_YEAR)
                            {
                                if (delimitedFileName[0] == yearTBx.Text)
                                {
                                    readExcelSheet(file.Name);
                                }
                            }
                            else
                            {
                                consoleOutputTxB.AppendText("Value is not in the year range. \n");
                            }
                            
                        }
                    }
                    else
                    {
                        consoleOutputTxB.Visible = true;
                        consoleBxLB.Visible = true;

                        consoleOutputTxB.AppendText("Value is not an integer. Please reenter the value \n");
                    }

                    break;
                case 3: //filter by semester
                    if (Int32.TryParse(yearTBx.Text.Trim(), out convertYear))
                    {
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (convertYear >= ZERO && convertYear <= FINAL_YEAR)
                            {
                                if (delimitedFileName[0] == yearTBx.Text)
                                {
                                    semesterSelected = semesterCmBx.SelectedItem.ToString();
                                    semesterSelected = semesterSelected.ToLower();
                                    if (delimitedFileName[1] == semesterSelected)
                                    {
                                        readExcelSheet(file.Name);
                                    }
                                }
                            }
                            else
                            {
                                consoleOutputTxB.AppendText("Value is not in the year range. \n");
                            }
                        }
                    }
                    else
                    {
                        consoleOutputTxB.Visible = true;
                        consoleBxLB.Visible = true;

                        consoleOutputTxB.AppendText("Value is not an integer. Please reenter the value \n");
                    }
                    break;
                case 4: //filter by course
                    if (Int32.TryParse(yearTBx.Text.Trim(), out convertYear))
                    {
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (convertYear >= ZERO && convertYear <= FINAL_YEAR)
                            {
                                if (delimitedFileName[0] == yearTBx.Text)
                                {
                                    semesterSelected = semesterCmBx.SelectedItem.ToString();
                                    semesterSelected = semesterSelected.ToLower();
                                    if (delimitedFileName[1] == semesterSelected)
                                    {
                                        if (delimitedFileName[2] == (string)courseCmBx.SelectedItem)
                                        {
                                            readExcelSheet(file.Name);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                consoleOutputTxB.AppendText("Value is not in the year range. \n");
                            }
                        }
                    }
                    else
                    {
                        consoleOutputTxB.Visible = true;
                        consoleBxLB.Visible = true;

                        consoleOutputTxB.AppendText("Value is not an integer. Please reenter the value \n");
                    }
                    break;
                case 5: //filter by section
                    if (Int32.TryParse(yearTBx.Text.Trim(), out convertYear))
                    {
                        foreach (FileInfo file in Files)
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (convertYear >= ZERO && convertYear <= FINAL_YEAR)
                            {
                                if (delimitedFileName[0] == yearTBx.Text)
                                {
                                    semesterSelected = semesterCmBx.SelectedItem.ToString();
                                    semesterSelected = semesterSelected.ToLower();
                                    if (delimitedFileName[1] == semesterSelected)
                                    {
                                        if (delimitedFileName[2] == (string)courseCmBx.SelectedItem)
                                        {
                                            if (Int32.TryParse(yearTBx.Text.Trim(), out convertSection))
                                            {
                                                if (delimitedFileName[3] == (string)sectionTBx.Text.Trim())
                                                {
                                                    readExcelSheet(file.Name);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                consoleOutputTxB.AppendText("Value is not in the year range. \n");
                            }
                        }
                    }
                    else
                    {
                        consoleOutputTxB.Visible = true;
                        consoleBxLB.Visible = true;

                        consoleOutputTxB.AppendText("Value is not an integer. Please reenter the value \n");
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
            if (noneRB.Checked == true)
            {
                yearTBx.Enabled = false;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionTBx.Enabled = false;
            }
            if (yearRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionTBx.Enabled = false;
            }
            if (semesterRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = false;
                sectionTBx.Enabled = false;
            }
            if (courseRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionTBx.Enabled = false;
            }
            if (sectionRB.Checked == true)
            {
                yearTBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionTBx.Enabled = true;
            }

        }

        private void pullDataIntoArray(String excelInfoGet, int indexAtPass)
        {
            switch(excelInfoGet)
            {
                case "objective1":
                    break;
                case "objective2":
                    break;
                case "objective3":
                    break;              
                default:
                    break;
            }
        }

    }
}

public class Node
{
    private Node objective;
    private int students;
    private int maxScore;
    private int actualScore;

    public Node()
    {
        objective = null;
        students = 0;
        maxScore = 0;
        actualScore = 0;
    }

    public void setStudents(int studentVal)
    {
        students = studentVal;
    }

    public int getStudents()
    {
        return students;
    }

    public void setMaxScore(int scoreMax)
    {
        maxScore = scoreMax;
    }

    public int getMaxScore()
    {
        return maxScore;
    }

    public void setActualScore(int studentActual)
    {
        actualScore = studentActual;
    }

    public int getActualScore()
    {
        return actualScore;
    }
}