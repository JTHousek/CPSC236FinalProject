using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAssessmentIntegration
{
    public partial class ProcessedWindow : Form
    {
        public ProcessedWindow()
        {
            InitializeComponent();
        }

        private void OkBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void displayObjectives(objNodeLL objectiveList)
        {
            double percentage = 0.0;
            objNode currNode = new objNode();
            currNode = objectiveList.getHead();

            while (currNode != null)
            {
                outputObjLBx.Items.Add(currNode.getObjective());
                outputNumLBx.Items.Add(currNode.getStudents().ToString());
                percentage = (currNode.getActualScore() / currNode.getMaxScore()) * 100.00;
                outputPercLBx.Items.Add(percentage.ToString("F2"));

                currNode = currNode.getNextObjNode();
            }

        }

        //a method to load in requested excelsheets based on their path in the file system
        public void readExcelSheet(string sheetPath, int sheetNum, objNodeLL objectiveList)
        {
            int rowCount = 0;

            //open excelApp and create the new application
            Excel.Application excelApp;
            excelApp = new Excel.Application();
            //workbook 
            Excel.Workbook excelWorkbook;
            //worksheet 
            Excel.Worksheet excelWorksheet;
            //range variable
            Excel.Range range;

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
                if (range.Cells.Rows.Count != 0)
                {

                    //Number of rows
                    Console.WriteLine("Number of Rows: " + range.Rows.Count);

                    //Number of columns
                    Console.WriteLine("Rumber of Columns: " + range.Columns.Count);

                    if (sheetNum == 1) //if this is the first sheet read in
                    {
                        objNode firstObj = new objNode();
                        rowCount = 1;

                        firstObj.setObjective(range.Cells[rowCount, 1].Value2.ToString());
                        firstObj.setStudents(range.Cells[rowCount, 2].Value2);
                        firstObj.setMaxScore(range.Cells[rowCount, 3].Value2);
                        firstObj.setActualScore(range.Cells[rowCount, 4].Value2);

                        objectiveList.setHead(firstObj);
                        objectiveList.setTail(firstObj); //there is only one node in the list
                    }

                    for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                    {
                        if (sheetNum == 1 && rowCount == 1) //if this is the first sheet, skip the first row which is done above
                        {
                            continue; //skip iteration
                        }

                        bool foundFlag = false; //flags if there's a duplicate obj
                        objNode newObj = new objNode();
                        objNode currNode = objectiveList.getHead();

                        newObj.setObjective(range.Cells[rowCount, 1].Value2.ToString());
                        newObj.setStudents(range.Cells[rowCount, 2].Value2);
                        newObj.setMaxScore(range.Cells[rowCount, 3].Value2);
                        newObj.setActualScore(range.Cells[rowCount, 4].Value2);

                        while (currNode != null)
                        {
                            if (currNode.getObjective() == newObj.getObjective())
                            {
                                currNode.setStudents(currNode.getStudents() + newObj.getStudents());
                                currNode.setMaxScore(currNode.getMaxScore() + newObj.getMaxScore());
                                currNode.setActualScore(currNode.getActualScore() + newObj.getActualScore());
                                foundFlag = true;
                                break; //the node was found, break the search
                            }
                            currNode = currNode.getNextObjNode();
                        }

                        if (!foundFlag)
                        {
                            objectiveList.getTail().setNextObjNode(newObj);
                            objectiveList.setTail(newObj); //set a new tail for the list 
                        }
                    }

                    //close
                    excelWorkbook.Close(true, null, null);
                    excelApp.Quit();
                }
                else
                {
                    MessageBox.Show("empty spreadsheet");
                }

            }
            catch (FileNotFoundException ex)
            {
                //consoleOutputTxB.AppendText("ERROR: FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }
        }
    }
}
