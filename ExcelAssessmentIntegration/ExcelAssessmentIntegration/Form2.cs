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

        //a method to load in requested excelsheets based on their path in the file system
        public void readExcelSheet(string sheetPath)
        {
            String objStr = "";
            double numStudents = 0;
            double maxScore = 0;
            double actualScore = 0;
            double percentage = 0.0;
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

                //Number of rows
                Console.WriteLine("Number of Rows: " + range.Rows.Count);

                //Number of columns
                Console.WriteLine("Rumber of Columns: " + range.Columns.Count);

                for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                {
                    objStr = range.Cells[rowCount, 1].Value2.ToString();
                    outputObjLBx.Items.Add(objStr);
                    numStudents = range.Cells[rowCount, 2].Value2;
                    outputNumLBx.Items.Add(numStudents.ToString());
                    maxScore = range.Cells[rowCount, 3].Value2;
                    actualScore = range.Cells[rowCount, 4].Value2;
                    percentage = (actualScore / maxScore) * 100.00;
                    outputPercLBx.Items.Add(percentage.ToString("F2"));



                    /*switch(objStr)
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
                    }*/


                }
                //MessageBox.Show(obj1.getStudents().ToString());

                //close
                excelWorkbook.Close(true, null, null);
                excelApp.Quit();


            }
            catch (FileNotFoundException ex)
            {
                //consoleOutputTxB.AppendText("ERROR: FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }
        }
    }
}
