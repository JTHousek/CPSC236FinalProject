//James Houseknecht, Cain Kaltenbaugh, Ethan Mooney
//jth1012, cxk1047, eam1020
//ExcelAssessmentIntegration
//Form1.cs
//Start Date: October 1, 2019
//End Date: December 1, 2019

using System;
using System.IO;
using System.Xml.Linq;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAssessmentIntegration
{
    public partial class ExcelIntegrationAssessmentWindow : Form
    {
        DirectoryInfo dataFilesDir = new DirectoryInfo("..\\dataSheets\\"); //specifies the directory that the dataFiles are in
        objNodeLL objectiveList = new objNodeLL();                          //a linked list to house the super objectives
        public ExcelIntegrationAssessmentWindow()
        {
            InitializeComponent();
        }

        //this method fires when the form is started
        private void formOnLoad(object sender, EventArgs e)
        {
            loadAvailableFiles();   //load in all the available files to the combo box
            criteriaXMLLoad();      //load the XML that specifies the criteria for the filters
            objectiveXMLLoad();     //load the objectives and the map for course objectives
        }

        //loads in the available files in dataFiles and checks them against the filter criteria
        private void loadAvailableFiles()
        {
            const int FILENAMEDELIMITEDLENGTH = 4; //makes sure there are 4 parts to the file name to avoid nulls
            FileInfo[] Files;                      //holds all the files in the dataFiles Directory
            string semesterSelected;               //needed to hold the semester selected as all lowercase when checking it against the file name version
            String[] delimitedFileName;            //holds the file name when broken into 4 parts

            semesterSelected = "";                 //sets default value for semester selected in case if never files
            Files = dataFilesDir.GetFiles("*.xlsx");
            filesLBx.Items.Clear();                //clears the list box when refiltering so that no files that don't
                                                   //fall under the filter stay in the box

            foreach (FileInfo file in Files)       //for each file in the directory
            {
                if (!selectedFilesLBx.Items.Contains(file.Name)) //make sure if already selected it doesn't go into the files box
                {
                    delimitedFileName = file.Name.Split('_');    //delimit the name based on the '_' character, making it 4 parts
                    if (delimitedFileName.Length == FILENAMEDELIMITEDLENGTH) { //make sure that at minimum, the first 4 cells of delmitedFileName are filled
                        if (yearCmBx.SelectedItem == null || delimitedFileName[0] == yearCmBx.SelectedItem.ToString()) //checks if the filter criteria is met for year
                        {                                                                                              //or if no year is selected at all
                            if (semesterCmBx.SelectedItem != null) //as long as a semester is selected
                            {
                                semesterSelected = semesterCmBx.SelectedItem.ToString().Trim();
                                semesterSelected = semesterSelected.ToLower();              //change the criteria to all lowercase to match file names
                            }
                            if (semesterCmBx.SelectedItem == null || delimitedFileName[1] == semesterSelected) //same as year for semester
                            {
                                if (courseCmBx.SelectedItem == null || delimitedFileName[2] == courseCmBx.SelectedItem.ToString().Trim()) //same as year for course
                                {
                                    if (sectionCmBx.SelectedItem == null || delimitedFileName[3].Split('.')[0] == sectionCmBx.SelectedItem.ToString()) //same as year for section
                                    {
                                        filesLBx.Items.Add(file.Name); //add the file to the displayed file box if it meets all criteria
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
            XElement xmlDoc = null; //an xelement to house the xml file
            int maxYear = 0;        //needed to house the max year when creating combo box criteria
            int maxSection = 0;     //same as max year for section

            try
            {
                xmlDoc = XElement.Load("..\\criteria.xml"); //load the criteria xml file in

                //read in each set of criteria into a variable 
                var year = (from elem in xmlDoc.Elements("maxYear").Elements("year") select elem.Value);
                var semesters = (from elem in xmlDoc.Elements("semesters").Elements("semester") select elem.Value);
                var courses = (from elem in xmlDoc.Elements("courses").Elements("course") select elem.Value);
                var section = (from elem in xmlDoc.Elements("maxSection").Elements("section") select elem.Value);
                //since year and section will only specify the maximum, it parses that max to an integer
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
            catch (FileNotFoundException ex) //catch if the xml file is not found
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

        //load in the objective map xml
        private void objectiveXMLLoad()
        {
            char objChar;                   //subobjective name character variable
            string objName;                 //full objective name
            XElement xmlDoc = null;         //xelement to load the objective map in
            //integers that record the number subobjectives per super objective
            int psctFieldsInt = 0;         
            int cisFieldsInt = 0;
            int eprFieldsInt = 0;
            int dpwcFieldsInt = 0;
            bool firstObjectiveSet = true;  //a flag to verify the head of the linked list is set

            try
            {
                xmlDoc = XElement.Load("..\\objectives.xml"); //load the objective map xml file in
                //read in the fields attribute for the superobjectives to record how many subobjectives there are 
                var psctFields = (from elem in xmlDoc.Elements("PSCT").Attributes("fields") select elem.Value);
                var cisFields = (from elem in xmlDoc.Elements("CIS").Attributes("fields") select elem.Value);
                var eprFields = (from elem in xmlDoc.Elements("EPR").Attributes("fields") select elem.Value);
                var dpwcFields = (from elem in xmlDoc.Elements("DPWC").Attributes("fields") select elem.Value);
                //parse the number of subobjectives to an int
                Int32.TryParse(psctFields.ElementAt(0).ToString(), out psctFieldsInt);
                Int32.TryParse(cisFields.ElementAt(0).ToString(), out cisFieldsInt);
                Int32.TryParse(eprFields.ElementAt(0).ToString(), out eprFieldsInt);
                Int32.TryParse(dpwcFields.ElementAt(0).ToString(), out dpwcFieldsInt);

                //create a node for each subobjective of PSCT
                for (int i = 1; i <= psctFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);           //the subobjective name based on a char
                    objName = "PSCT" + objChar.ToString();  //create the subobjective name used in the xml file
                    newObjNode.setObjective("Problem Solving and Critical Thinking " + objChar.ToString()); //set the objective name for the node and output

                    //variable the houses the course objectives mapped to the subobjective
                    var psctCourseObjs = (from elem in xmlDoc.Elements("PSCT").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in psctCourseObjs)
                    {
                        newObjNode.addCourseObj(parse); //add the course objectives to the node
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    { 
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;          //make sure the head is not set again
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode); //set the current tail's next node to the current one
                        objectiveList.setTail(newObjNode);                  //set the tail of the list to the current node
                    }
                }

                //create a node for each subobjective of CIS
                for (int i = 1; i <= cisFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);           //the subobjective name based on a char
                    objName = "CIS" + objChar.ToString();   //create the subobjective name used in the xml file
                    newObjNode.setObjective("Communication and Interpersonal Skills " + objChar.ToString());

                    //variable the houses the course objectives mapped to the subobjective
                    var cisCourseObjs = (from elem in xmlDoc.Elements("CIS").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in cisCourseObjs)
                    {
                        newObjNode.addCourseObj(parse); //add the course objectives to the node
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    {
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;          //make sure the head is not set again
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode); //set the current tail's next node to the current one
                        objectiveList.setTail(newObjNode);                  //set the tail of the list to the current node
                    }
                }

                //create a node for each subobjective of EPR
                for (int i = 1; i <= eprFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);           //the subobjective name based on a char
                    objName = "EPR" + objChar.ToString();   //create the subobjective name used in the xml file
                    newObjNode.setObjective("Ethical and Professional Responsibilities " + objChar.ToString());

                    //variable the houses the course objectives mapped to the subobjective
                    var eprCourseObjs = (from elem in xmlDoc.Elements("EPR").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in eprCourseObjs)
                    {
                        newObjNode.addCourseObj(parse); //add the course objectives to the node
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    {
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;          //make sure the head is not set again
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode); //set the current tail's next node to the current one
                        objectiveList.setTail(newObjNode);                  //set the tail of the list to the current node
                    }
                }

                //create a node for each subobjective of DPWC
                for (int i = 1; i <= dpwcFieldsInt; i++)
                {
                    objNode newObjNode = new objNode(); //the node with all the info being added
                    objChar = (char)(i + 96);           //the subobjective name based on a char
                    objName = "DPWC" + objChar.ToString();  //create the subobjective name used in the xml file
                    newObjNode.setObjective("Degree Program Writing Competency " + objChar.ToString());

                    //variable the houses the course objectives mapped to the subobjective
                    var dpwcCourseObjs = (from elem in xmlDoc.Elements("DPWC").Elements(objName).Elements("courseObj") select elem.Value);
                    foreach (var parse in dpwcCourseObjs)
                    {
                        newObjNode.addCourseObj(parse); //add the course objectives to the node
                    }
                    if (firstObjectiveSet) //if this is the first objective overall
                    {
                        objectiveList.setHead(newObjNode);
                        objectiveList.setTail(newObjNode);
                        firstObjectiveSet = false;          //make sure the head is not set again
                    }
                    else
                    {
                        objectiveList.getTail().setNextObjNode(newObjNode); //set the current tail's next node to the current one
                        objectiveList.setTail(newObjNode);                  //set the tail of the list to the current node
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

        //load the output excel sheet with the current linked list
        public void displayObjectives()
        {
            objNode currNode;   //used to iterate through the list
            int rowNum = 1;     //used to iterate through the spreadsheet
        
            Excel.Application excelApp;     //open excelApp
            Excel.Workbook excelWorkbook;   //open a new workbook
            Excel.Worksheet excelWorksheet; //open a new worksheet

            currNode = objectiveList.getHead(); //set the current node to the head
            excelApp = new Excel.Application(); //create the new application
            excelApp.Visible = true;            //make the sheet visible
            excelWorkbook = excelApp.Workbooks.Add();
            excelWorksheet = excelWorkbook.Worksheets.get_Item(1);

            excelWorksheet.Cells[rowNum, 1] = "Objective Name";     //create the header row
            excelWorksheet.Cells[rowNum, 2] = "Weighted Average";
            rowNum++;                                               //go to the next row

            while (currNode != null) //while there are still nodes in the linked list
            {
                currNode.computeWeightedAverage();  //compute the node's weighted average
                excelWorksheet.Cells[rowNum, 1] = currNode.getObjective(); //set the first column of the row to the objective name
                if (currNode.getTotalStudents() == 0) //if there's no data to use
                {
                    excelWorksheet.Cells[rowNum, 2] = "-"; //just display a -
                }
                else //if there is data to use, display it in the second column
                {
                    excelWorksheet.Cells[rowNum, 2] = currNode.getWeightedAverage().ToString("F2");
                }
                rowNum++; //move to the next row
                currNode = currNode.getNextObjNode(); //move to the next node
            }

            excelWorksheet.Columns.AutoFit(); //autofit the rows to the size of the data

        }

        //a method to load in requested excelsheets based on their path in the file system
        public void readExcelSheet(string sheetPath)
        {
            int newStudents;       //holds the values read from the spreadsheet to parse to
            double newMaxScore;
            double newActualScore;
            int rowCount = 0; //needed to iterate through the spreadsheets
            objNode currNode; //needed to iterate through the linked list
            string[] delimitedFileName = sheetPath.Split('_'); //breaks the file name into pieces based on the '_' char
            string[] delimitedCourseObjName; //delimits the course objective name in the nodes

            Excel.Application excelApp;     //open excelApp
            Excel.Workbook excelWorkbook;   //open a new workbook
            Excel.Worksheet excelWorksheet; //open a new worksheet

            excelApp = new Excel.Application(); //create the new application
            
            Excel.Range range;  //range variable

            //get the current file path
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            //path to the workbook in the dataSheets folder
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
                    //for each row in the spreadsheet
                    for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                    {
                        currNode = objectiveList.getHead(); //set current node to the first node in the list

                        while (currNode != null) //while there are objective nodes left
                        {
                            foreach (string courseObj in currNode.getCourseObj()) //for each course objective mapped to the super objective
                            {
                                delimitedCourseObjName = courseObj.Split('-'); //delimit the objective into course and objective name
                                if (delimitedFileName[2] == delimitedCourseObjName[0] &&                            //if the objective and course
                                    range.Cells[rowCount, 1].Value2.ToString().Trim() == delimitedCourseObjName[1]) //name match
                                {
                                    int.TryParse(range.Cells[rowCount, 2].Value2.ToString(), out newStudents);      //parse the values to
                                    double.TryParse(range.Cells[rowCount, 3].Value2.ToString(), out newMaxScore);   //the node from the
                                    double.TryParse(range.Cells[rowCount, 4].Value2.ToString(), out newActualScore);//spreadsheet
                                    currNode.addStudents(newStudents);
                                    currNode.setTotalStudents(currNode.getTotalStudents() + newStudents); //add the new number of students to the total
                                    currNode.addAverage(newActualScore / newMaxScore);
                                }
                            }
                            currNode = currNode.getNextObjNode(); //get the next node in the list
                        }
                    }

                    //close
                    excelWorkbook.Close(true, null, null);
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    consoleOutputTxB.Visible = true;
                    consoleBxLB.Visible = true;
                    consoleOutputTxB.AppendText("ERROR: Exception: " + ex.GetType() + "\n");

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
                selectedFilesLBx.Items.Add(filesLBx.SelectedItem); //add the file to selected files
                filesLBx.Items.Remove(filesLBx.SelectedItem);      //remove the file from available files
            }
        }

        private void RemoveSheetBtn_Click(object sender, EventArgs e) //remove a sheet from the selected listbox
        {
            if (selectedFilesLBx.SelectedItem != null) //verify that something is selected
            {
                filesLBx.Items.Add(selectedFilesLBx.SelectedItem);              //add the files to available files
                selectedFilesLBx.Items.Remove(selectedFilesLBx.SelectedItem);   //remove the files from selected files
            }
        }

        //reloads the files in the files list box based on the current filter criteria
        private void FilterBtn_Click(object sender, EventArgs e)
        {
            loadAvailableFiles(); 
        }

        //the method tied to the Read Excel Sheets button that does most of the processing
        private void ReadExcelBtn_Click(object sender, EventArgs e)
        {
            //this needs to make sure there isn't one already open FUCK FUCK FUCK FUCK FUCK

            foreach (string item in selectedFilesLBx.Items) //for each file in the selected list of files
            {
                readExcelSheet(item); //read in the file to the linked list
            }

            displayObjectives(); //display the output at the end in a new excel sheet
        }

        //method which fires when cancel is clicked, closing the window
        private void CancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();   //close the form window
        }

        //moves all available files in the files list box to the selected list box
        private void AddAllBtn_Click(object sender, EventArgs e)
        {
            int numOfFiles = filesLBx.Items.Count; //records the number of files currently available
            for (int i = 0; i < numOfFiles; i++)
            {
                selectedFilesLBx.Items.Add(filesLBx.Items[0]); //add the file to selected files
                filesLBx.Items.Remove(filesLBx.Items[0]);      //remove the file from available files
            }
        }

        //moves all available files in the selected list box to the available list box
        private void RemoveAllBtn_Click(object sender, EventArgs e)
        {
            int numOfFiles = selectedFilesLBx.Items.Count; //records the number of files currently selected
            for (int i = 0; i < numOfFiles; i++)
            {
                filesLBx.Items.Add(selectedFilesLBx.Items[0]);              //add the file to available files
                selectedFilesLBx.Items.Remove(selectedFilesLBx.Items[0]);   //remove the file from selected files
            }
        }

        //method tied to the clear filter button that sets the filter criteria to null and reloads the available files
        private void clearFilterCriteriaBtn_Click(object sender, EventArgs e)
        {
            //clear each combo box
            yearCmBx.SelectedItem = null;
            semesterCmBx.SelectedItem = null;
            courseCmBx.SelectedItem = null;
            sectionCmBx.SelectedItem = null;

            //reload the files
            loadAvailableFiles();
        }
    }
}