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
            FileInfo[] Files = dataFilesDir.GetFiles("*.xlsx"); //Getting Text files
            int filterType = 1; //none by default for when the form first loads
            string semesterSelected = "";
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
                            if (yearCmBx.SelectedItem == null || delimitedFileName[0] == yearCmBx.SelectedItem.ToString()) //removes the xml insertion artifacting
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
                            if (yearCmBx.SelectedItem == null || delimitedFileName[0] == yearCmBx.SelectedItem.ToString())
                            {
                                if (semesterCmBx.SelectedItem != null) //as long as a semester is selected
                                {
                                    semesterSelected = semesterCmBx.SelectedItem.ToString().Split(',')[1].Trim().Split(']')[0];
                                    semesterSelected = semesterSelected.ToLower();
                                }
                                if (semesterCmBx.SelectedItem == null || delimitedFileName[1] == semesterSelected)
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
                    break;
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
                year.AddRange((from elem in element.Descendants("year") select new KeyValuePair<int, string>((int) elem.Attribute("key"), (string) elem.Attribute("value"))));
                semesters.AddRange((from elem in element.Descendants("semester") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                courses.AddRange((from elem in element.Descendants("course") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
                section.AddRange((from elem in element.Descendants("section") select new KeyValuePair<int, string>((int)elem.Attribute("key"), (string)elem.Attribute("value"))));
            }
            catch (FileNotFoundException ex)
            {
                consoleOutputTxB.AppendText("ERROR: XML FILE NOT READ, Exception: " + ex.GetType() + "\n");
            }

            Int32.TryParse(year.ElementAt(0).Value, out maxYear);
            Int32.TryParse(section.ElementAt(0).Value, out maxSection);
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
            //this needs to make sure there isn't one already open
            ProcessedWindow processWindow = new ProcessedWindow();
            int sheetNum = 1; //used to verify which sheet is first to set LL head
            objNodeLL objectiveList = new objNodeLL();

            foreach(string item in selectedFilesLBx.Items)
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

        //which boxes are active and have a selected element is handled when changing criteria
        private void filterCriteriaGrpBx_CheckedChanged(object sender, EventArgs e)
        {
            if (noneRB.Checked == true)
            {
                yearCmBx.Enabled = false;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;

                yearCmBx.SelectedItem = null;
                semesterCmBx.SelectedItem = null;
                courseCmBx.SelectedItem = null;
                sectionCmBx.SelectedItem = null;
            }
            if (yearRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = false;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;

                semesterCmBx.SelectedItem = null;
                courseCmBx.SelectedItem = null;
                sectionCmBx.SelectedItem = null;
            }
            if (semesterRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = false;
                sectionCmBx.Enabled = false;

                courseCmBx.SelectedItem = null;
                sectionCmBx.SelectedItem = null;
            }
            if (courseRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionCmBx.Enabled = false;

                sectionCmBx.SelectedItem = null;
            }
            if (sectionRB.Checked == true)
            {
                yearCmBx.Enabled = true;
                semesterCmBx.Enabled = true;
                courseCmBx.Enabled = true;
                sectionCmBx.Enabled = true;
            }

        }
    }
}