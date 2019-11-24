//James Houseknecht, Cain Kaltenbaugh, Ethan Mooney
//jth1012, cxk1047, eam1020
//ExcelAssessmentIntegration
//objNode.cs
//Start Date: October 1, 2019
//End Date: December 1, 2019

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAssessmentIntegration
{
    public class objNode
    {
        private string objective;       //the name of the super objective
        private int totalStudents;      //the total students from all course objectives
        private double weightedAverage; //the weightedAverage of all course objectives
        private objNode nextObjNode;    //the next objective node in the linked list

        private List<int> students = new List<int>();           //list of number of students per course objective
        private List<double> average = new List<double>();      //list of averages per course objective (actual/max)
        private List<string> courseObjs = new List<string>();   //list of course objectives specified by an XML file read on startup

        //objective node constructor
        public objNode()
        {
            objective = null;
            totalStudents = 0;
            weightedAverage = 0.0;
            nextObjNode = null;
        }

        //computes the weighted average of the whole super objective
        //from taking the number of students and the averages for each
        //course objective
        public void computeWeightedAverage()
        {
            for (int i = 0; i < students.Count; i++)
            {
                weightedAverage = weightedAverage + (students[i] * average[i]);
            }
            weightedAverage = weightedAverage / totalStudents;
        }

        //returns the weighted average
        public double getWeightedAverage()
        {
            return weightedAverage;
        }

        //sets the total amount of students in the whole super objective
        public void setTotalStudents(int totalStudents)
        {
            this.totalStudents = totalStudents;
        }

        //returns the total number of students
        public int getTotalStudents()
        {
            return totalStudents;
        }

        //sets the name of the super objective
        public void setObjective(string objective)
        {
            this.objective = objective;
        }

        //returns the super objective name
        public string getObjective()
        {
            return objective;
        }

        //adds a new course objectives' number of students to the students list
        public void addStudents(int newStudents)
        {
            students.Add(newStudents);
        }

        //returns the list of course objectives' number of students
        public List<int> getStudents()
        {
            return students;
        }

        //adds a new course objectives' averages to the averages list
        public void addAverage(double newAverage)
        {
            average.Add(newAverage);
        }

        //returns the list of course objectives' averages
        public List<double> getAverage()
        {
            return average;
        }

        //adds the course objectives attributed to the super objective through the XML sheet
        public void addCourseObj(string newCourseObj)
        {
            courseObjs.Add(newCourseObj);
        }

        //returns the list of the course objectives tied to the super objective
        public List<string> getCourseObj()
        {
            return courseObjs;
        }

        //sets the next node in the linked list
        public void setNextObjNode(objNode nextObjNode)
        {
            this.nextObjNode = nextObjNode;
        }

        //gets the next node in the linked list
        public objNode getNextObjNode()
        {
            return nextObjNode;
        }
    }
}
