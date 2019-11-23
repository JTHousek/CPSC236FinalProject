using System;
using System.IO;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Collections;
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
        objNodeLL objectiveList = new objNodeLL();
        public ExcelIntegrationAssessmentWindow()
        {
            InitializeComponent();
        }

        private void formOnLoad(object sender, EventArgs e)
        {
            loadAvailableFiles();
            criteriaXMLLoad();
            objectiveXMLLoad();
        }

        private void loadAvailableFiles()
        {
            const int FILENAMEDELIMITEDLENGTH = 4;
            FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
            string semesterSelected = "";
            String[] delimitedFileName;
            int filesCount = dataFilesDir.GetFiles().Length;
            filesLBx.Items.Clear();

            foreach (FileInfo file in Files)
            {
                if (!selectedFilesLBx.Items.Contains(file.Name)) //make sure if already selected it doesnt go into the files box
                {
                    delimitedFileName = file.Name.Split('_');
                    if (delimitedFileName.Length == FILENAMEDELIMITEDLENGTH) {
                        if (yearCmBx.SelectedItem == null || delimitedFileName[0] == yearCmBx.SelectedItem.ToString())
                        {
                            if (semesterCmBx.SelectedItem != null) //as long as a semester is selected
                            {
                                semesterSelected = semesterCmBx.SelectedItem.ToString().Trim();
                                semesterSelected = semesterSelected.ToLower();
                            }
                            if (semesterCmBx.SelectedItem == null || delimitedFileName[1] == semesterSelected)
                            {
                                if (courseCmBx.SelectedItem == null || delimitedFileName[2] == courseCmBx.SelectedItem.ToString().Trim())
                                {
                                    if (sectionCmBx.SelectedItem == null || delimitedFileName[3].Split('.')[0] == sectionCmBx.SelectedItem.ToString())
                                    {
                                        filesLBx.Items.Add(file.Name);
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }

        //a method which loads startup.xml which houses the contents of each filter dropdown
        private void criteriaXMLLoad()
        {
            XElement xmlDoc = null;
            int maxYear = 0;
            int maxSection = 0;
            try
            {
                xmlDoc = XElement.Load("..\\criteria.xml");
                var year = (from elem in xmlDoc.Elements("maxYear").Elements("year") select elem.Value);
                var semesters = (from elem in xmlDoc.Elements("semesters").Elements("semester") select elem.Value);
                var courses = (from elem in xmlDoc.Elements("courses").Elements("course") select elem.Value);
                var section = (from elem in xmlDoc.Elements("maxSection").Elements("section") select elem.Value);
                Int32.TryParse(year.ElementAt(0).ToString(), out maxYear);
                Int32.TryParse(section.ElementAt(0).ToString(), out maxSection);

                //load the semesters combo box based on the semesters provided
                foreach (string parse in semesters)
                {
                    semesterCmBx.Items.Add(parse);
                }

                //load the courses combo box based on the courses provided
                foreach (string parse in courses)
                {
                    courseCmBx.Items.Add(parse);
                }
            }
            catch (FileNotFoundException ex)
            {
                consoleOutputTxB.Visible = true;
                consoleBxLB.Visible = true;
                consoleOutputTxB.AppendText("ERROR: XML FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }

            //load the years combo box based on the max year provided
            for (int i = 2000; i <= maxYear; i++)
            {
                yearCmBx.Items.Add((i - 2000).ToString("00"));
            }

            //load the sections combo box based on the max section provided
            for (int i = 1; i <= maxSection; i++)
            {
                sectionCmBx.Items.Add((i).ToString());
            }

            //makes sure selected items are blank at initialization
            yearCmBx.SelectedItem = null;
            semesterCmBx.SelectedItem = null;
            courseCmBx.SelectedItem = null;
            semesterCmBx.SelectedItem = null;
        }

        private void objectiveXMLLoad()
        {
            XElement xmlDoc = null;
            int psctFieldsInt = 0;
            int cisFieldsInt = 0;
            int eprFieldsInt = 0;
            int dpwcFieldsInt = 0;
            char objChar;
            string objName;
            bool firstObjectiveSet = true;

            try
            {
                xmlDoc = XElement.Load("..\\objectives.xml");
                var psctFields = (from elem in xmlDoc.Elements("PSCT").Attributes("fields") select elem.Value);
                var cisFields = (from elem in xmlDoc.Elements("CIS").Attributes("fields") select elem.Value);
                var eprFields = (from elem in xmlDoc.Elements("EPR").Attributes("fields") select elem.Value);
                var dpwcFields = (from elem in xmlDoc.Elements("DPWC").Attributes("fields") select elem.Value);
                Int32.TryParse(psctFields.ElementAt(0).ToString(), out psctFieldsInt);
                Int32.TryParse(cisFields.ElementAt(0).ToString(), out cisFieldsInt);
                Int32.TryParse(eprFields.ElementAt(0).ToString(), out eprFieldsInt);
                Int32.TryParse(dpwcFields.ElementAt(0).ToString(), out dpwcFieldsInt);



                for (int i = 1; i <= psctFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);
                    objName = "PSCT" + objChar.ToString();
                    newObjNode.setObjective("Problem Solving and Critical Thinking " + objChar.ToString());

                    var psctCourseObjs = (from elem in xmlDoc.Elements("PSCT").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in psctCourseObjs)
                    {
                        newObjNode.addCourseObj(parse);
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    { 
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode);
                        objectiveList.setTail(newObjNode);
                    }
                }


                for (int i = 1; i <= cisFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);
                    objName = "CIS" + objChar.ToString();
                    newObjNode.setObjective("Communication and Interpersonal Skills " + objChar.ToString());

                    var cisCourseObjs = (from elem in xmlDoc.Elements("CIS").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in cisCourseObjs)
                    {
                        newObjNode.addCourseObj(parse);
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    {
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode);
                        objectiveList.setTail(newObjNode);
                    }
                }

                for (int i = 1; i <= eprFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);
                    objName = "EPR" + objChar.ToString();
                    newObjNode.setObjective("Ethical and Professional Responsibilities " + objChar.ToString());

                    var eprCourseObjs = (from elem in xmlDoc.Elements("EPR").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in eprCourseObjs)
                    {
                        newObjNode.addCourseObj(parse);
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    {
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode);
                        objectiveList.setTail(newObjNode);
                    }
                }

                for (int i = 1; i <= dpwcFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);
                    objName = "DPWC" + objChar.ToString();
                    newObjNode.setObjective("Degree Program Writing Competency " + objChar.ToString());

                    var dpwcCourseObjs = (from elem in xmlDoc.Elements("DPWC").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in dpwcCourseObjs)
                    {
                        newObjNode.addCourseObj(parse);
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    {
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode);
                        objectiveList.setTail(newObjNode);
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                consoleOutputTxB.Visible = true;
                consoleBxLB.Visible = true;
                consoleOutputTxB.AppendText("ERROR: XML FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }
        }

        public void displayObjectives()
        {
            int rowNum = 1;
            double percentage = 0.0;
            //open excelApp and create the new application
            Excel.Application excelApp;
            excelApp = new Excel.Application();
            excelApp.Visible = true;
            //workbook 
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            //worksheet 
            Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets.get_Item(1);
            //y
            objNode currNode = objectiveList.getHead();

            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            //change
            string workbookPath = path + "\\bin\\outputSheets\\" + "testOutput";

            excelWorksheet.Cells[rowNum, 1] = "Objective Name";
            excelWorksheet.Cells[rowNum, 2] = "Number of Students";
            excelWorksheet.Cells[rowNum, 3] = "Average";
            rowNum++;

            while (currNode != null)
            {
                excelWorksheet.Cells[rowNum, 1] = currNode.getObjective();
                excelWorksheet.Cells[rowNum, 2] = currNode.getStudents().ToString(); 
                percentage = (currNode.getActualScore() / currNode.getMaxScore()) * 100.00;
                excelWorksheet.Cells[rowNum, 3] = (percentage.ToString("F2"));
                rowNum++;
                currNode = currNode.getNextObjNode();
            }

            excelWorksheet.Columns.AutoFit();

        }

        //a method to load in requested excelsheets based on their path in the file system
        public void readExcelSheet(string sheetPath)
        {
            int rowCount = 0;
            string[] delimitedFileName = sheetPath.Split('_');
            string[] delimitedCourseObjName;

            //open excelApp and create the new application
            Excel.Application excelApp;
            excelApp = new Excel.Application();
            //workbook 
            Excel.Workbook excelWorkbook;
            //worksheet 
            Excel.Worksheet excelWorksheet;
            //range variable
            Excel.Range range;

            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            //change
            string workbookPath = path + "\\bin\\dataSheets\\" + sheetPath;

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
                try
                {
                    for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                    {
                        objNode currNode = objectiveList.getHead();

                        while (currNode != null)
                        {
                            foreach (string courseObj in currNode.getCourseObj())
                            {
                                delimitedCourseObjName = courseObj.Split('-');
                                if (delimitedFileName[2] == delimitedCourseObjName[0] && 
                                    range.Cells[rowCount, 1].Value2.ToString().Trim() == delimitedCourseObjName[1])
                                {
                                    currNode.setStudents(currNode.getStudents() + range.Cells[rowCount, 2].Value2);
                                    currNode.setMaxScore(currNode.getMaxScore() + range.Cells[rowCount, 3].Value2);
                                    currNode.setActualScore(currNode.getActualScore() + range.Cells[rowCount, 4].Value2);
                                }
                            }
                            currNode = currNode.getNextObjNode();
                        }
                    }

                    //close
                    excelWorkbook.Close(true, null, null);
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    excelWorkbook.Close(true, null, null);
                    excelApp.Quit();
                }
            }
            catch (FileNotFoundException ex)
            {
                consoleOutputTxB.Visible = true;
                consoleBxLB.Visible = true;
                consoleOutputTxB.AppendText("ERROR: FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }
        }
    

        private void AddSheetBtn_Click(object sender, EventArgs e) //add a sheet to the selected listbox
        {
            if (filesLBx.SelectedItem != null) //verify that something is selected 
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
            //this needs to make sure there isn't one already open
            int sheetNum = 1; //used to verify which sheet is first to set LL head

            foreach (string item in selectedFilesLBx.Items)
            {
                readExcelSheet(item);
                sheetNum++; //increment number of sheets
            }

            displayObjectives();
        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AddAllBtn_Click(object sender, EventArgs e)
        {
            int numOfFiles = filesLBx.Items.Count;
            for (int i = 0; i < numOfFiles; i++)
            {
                selectedFilesLBx.Items.Add(filesLBx.Items[0]);
                filesLBx.Items.Remove(filesLBx.Items[0]);
            }
        }

        private void RemoveAllBtn_Click(object sender, EventArgs e)
        {
            int numOfFiles = selectedFilesLBx.Items.Count;
            for (int i = 0; i < numOfFiles; i++)
            {
                filesLBx.Items.Add(selectedFilesLBx.Items[0]);
                selectedFilesLBx.Items.Remove(selectedFilesLBx.Items[0]);
            }
        }

        private void clearFilterCriteriaBtn_Click(object sender, EventArgs e)
        {
            FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
            const int FILENAMEDELIMITEDLENGTH = 4;
            String[] delimitedFileName;

            yearCmBx.SelectedItem = null;
            semesterCmBx.SelectedItem = null;
            courseCmBx.SelectedItem = null;
            sectionCmBx.SelectedItem = null;

            foreach (FileInfo file in Files)
            {
                if (!selectedFilesLBx.Items.Contains(file.Name) && !filesLBx.Items.Contains(file.Name)) //make sure if already selected it doesnt go into the files box
                {
                    delimitedFileName = file.Name.Split('_');
                    if (delimitedFileName.Length == FILENAMEDELIMITEDLENGTH)
                    {
                        filesLBx.Items.Add(file.Name);
                    }
                }
            }
        }
    }
}