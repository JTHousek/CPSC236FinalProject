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

namespace ExcelAssessmentIntegration
{
    public partial class ExcelIntegrationAssessmentWindow : Form
    {
        System.IO.DirectoryInfo dataFilesDir = new System.IO.DirectoryInfo("..\\dataSheets\\");
        public ExcelIntegrationAssessmentWindow()
        {
            InitializeComponent();
        }

        private void formOnLoad(object sender, EventArgs e)
        {
            loadAvailableFiles();
            criteriaXMLLoad();
            //objectiveXMLLoad();
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
            ProcessedWindow processWindow = new ProcessedWindow();
            int sheetNum = 1; //used to verify which sheet is first to set LL head
            objNodeLL objectiveList = new objNodeLL();

            foreach (string item in selectedFilesLBx.Items)
            {
                processWindow.readExcelSheet(item, sheetNum, objectiveList);
                sheetNum++; //increment number of sheets
            }

            processWindow.displayObjectives(objectiveList);
            processWindow.ShowDialog();
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