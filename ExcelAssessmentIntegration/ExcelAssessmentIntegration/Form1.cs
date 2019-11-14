using System;
using System.IO;
using System.Xml.Linq;
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

        private void formOnLoad(object sender, EventArgs e)
        {
            loadAvailableFiles();
            startupXMLLoad();
        }

        private void loadAvailableFiles()
        {
            FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
            int filterType = 1; //none by default for when the form first loads
            string semesterSelected;
            String[] delimitedFileName;
            int filesCount = dataFilesDir.GetFiles().Length;
            filesLBx.Items.Clear();
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

            switch (filterType)
            {
                case 1: //filter by nothing    
                    foreach (FileInfo file in Files)
                    {
                        if (!selectedFilesLBx.Items.Contains(file.Name)) //if file is already selected, don't add
                        { 
                            filesLBx.Items.Add(file.Name);
                        }
                    }
                    break;

                case 2: //filter by year 
                    foreach (FileInfo file in Files)
                    {
                        if (!selectedFilesLBx.Items.Contains(file.Name))
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[0] == yearCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0]) //removes the xml insertion artifacting
                            {
                                filesLBx.Items.Add(file.Name);
                            }
                        }
                    }
                    break;

                case 3: //filter by semester
                    foreach (FileInfo file in Files)
                    {
                        if (!selectedFilesLBx.Items.Contains(file.Name)) { 
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[0] == yearCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
                            {
                                semesterSelected = semesterCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0];
                                semesterSelected = semesterSelected.ToLower();
                                if (delimitedFileName[1] == semesterSelected)
                                {
                                    filesLBx.Items.Add(file.Name);
                                }
                            }
                        }
                    }
                    break;

                case 4: //filter by course
                    foreach (FileInfo file in Files)
                    {
                        if (!selectedFilesLBx.Items.Contains(file.Name)) { 
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[0] == yearCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
                            {
                                semesterSelected = semesterCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0];
                                semesterSelected = semesterSelected.ToLower();
                                if (delimitedFileName[1] == semesterSelected)
                                {
                                    if (delimitedFileName[2] == courseCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
                                    {
                                        filesLBx.Items.Add(file.Name);
                                    }
                                }
                            }
                        }
                    }
                    break;

                case 5: //filter by section
                    foreach (FileInfo file in Files)
                    {
                        if (!selectedFilesLBx.Items.Contains(file.Name))
                        {
                            delimitedFileName = file.Name.Split('_');
                            if (delimitedFileName[0] == yearCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
                            {
                                semesterSelected = semesterCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0];
                                semesterSelected = semesterSelected.ToLower();
                                if (delimitedFileName[1] == semesterSelected)
                                {
                                    if (delimitedFileName[2] == courseCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
                                    {
                                        if (delimitedFileName[3].Split('.')[0] == sectionCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
                                        {
                                            filesLBx.Items.Add(file.Name);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;
            }

        }

        //a method which loads startup.xml which houses the contents of each filter dropdown
        private void startupXMLLoad()
        {
            XElement element = null;
            List<KeyValuePair<int, string>> years = new List<KeyValuePair<int, string>>();
            List<KeyValuePair<int, string>> semesters = new List<KeyValuePair<int, string>>();
            List<KeyValuePair<int, string>> courses = new List<KeyValuePair<int, string>>();
            List<KeyValuePair<int, string>> sections = new List<KeyValuePair<int, string>>();
            try
            {
                element = XElement.Load("..\\startup.xml");
                years.AddRange((from elem in element.Descendants("year") select new KeyValuePair<int, string>((int) elem.Attribute("key"), (string) elem.Attribute("value"))));
                semesters.AddRange((from elem in element.Descendants("semester") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                courses.AddRange((from elem in element.Descendants("course") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                sections.AddRange((from elem in element.Descendants("section") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
            }
            catch (FileNotFoundException ex)
            {
                consoleOutputTxB.AppendText("ERROR: XML FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }
            yearCmBx.DataSource = years;
            yearCmBx.ValueMember = "Key";
            yearCmBx.DisplayMember = "Value";
            semesterCmBx.DataSource = semesters;
            semesterCmBx.ValueMember = "Key";
            semesterCmBx.DisplayMember = "Value";
            courseCmBx.DataSource = courses;
            courseCmBx.ValueMember = "Key";
            courseCmBx.DisplayMember = "Value";
            sectionCmBx.DataSource = sections;
            sectionCmBx.ValueMember = "Key";
            sectionCmBx.DisplayMember = "Value";
        }
        
        //a method to load in requested excelsheets based on their path in the file system
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
        private void AddSheetBtn_Click(object sender, EventArgs e) //add a sheet to the selected listbox
        {
            if(filesLBx.SelectedItem != null) //verify that something is selected 
            {
                selectedFilesLBx.Items.Add(filesLBx.SelectedItem);
                filesLBx.Items.Remove(filesLBx.SelectedItem);
            }
        }

        private void RemoveSheetBtn_Click(object sender, EventArgs e) //remove a sheet from the selected listbox
        {
            if (selectedFilesLBx.SelectedItem != null) //verify that something is selected
            {
                filesLBx.Items.Add(selectedFilesLBx.SelectedItem);
                selectedFilesLBx.Items.Remove(selectedFilesLBx.SelectedItem);
            }
        }

        private void FilterBtn_Click(object sender, EventArgs e)
        {
            loadAvailableFiles();
        }

        private void ReadExcelBtn_Click(object sender, EventArgs e)
        {
            foreach(string item in selectedFilesLBx.Items)
            {
                readExcelSheet(item);
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
                yearCmBx.Enabled = false;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;
            }
            if (yearRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;
            }
            if (semesterRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;
            }
            if (courseRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionCmBx.Enabled = false;
            }
            if (sectionRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionCmBx.Enabled = true;
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