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
            startupXMLLoad();
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
                                semesterSelected = semesterCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0];
                                semesterSelected = semesterSelected.ToLower();
                            }
                            if (semesterCmBx.SelectedItem == null || delimitedFileName[1] == semesterSelected)
                            {
                                if (courseCmBx.SelectedItem == null || delimitedFileName[2] == courseCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0])
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
        private void startupXMLLoad()
        {
            XElement element = null;
            int maxYear = 0;
            int maxSection = 0;
            List<KeyValuePair<int, string>> year = new List<KeyValuePair<int, string>>();
            List<KeyValuePair<int, string>> semesters = new List<KeyValuePair<int, string>>();
            List<KeyValuePair<int, string>> courses = new List<KeyValuePair<int, string>>();
            List<KeyValuePair<int, string>> section = new List<KeyValuePair<int, string>>();
            try
            {
                element = XElement.Load("..\\startup.xml");
                year.AddRange((from elem in element.Descendants("year") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                semesters.AddRange((from elem in element.Descendants("semester") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                courses.AddRange((from elem in element.Descendants("course") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                section.AddRange((from elem in element.Descendants("section") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                Int32.TryParse(year.ElementAt(0).Value, out maxYear);
                Int32.TryParse(section.ElementAt(0).Value, out maxSection);
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
            //load the semesters combo box based on the semesters provided
            semesterCmBx.DataSource = semesters;
            semesterCmBx.ValueMember = "Key";
            semesterCmBx.DisplayMember = "Value";
            //load the courses combo box based on the courses provided
            courseCmBx.DataSource = courses;
            courseCmBx.ValueMember = "Key";
            courseCmBx.DisplayMember = "Value";
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
            semesterCmBx.SelectedItem = null;

            foreach (FileInfo file in Files)
            {
                if (!selectedFilesLBx.Items.Contains(file.Name)) //make sure if already selected it doesnt go into the files box
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