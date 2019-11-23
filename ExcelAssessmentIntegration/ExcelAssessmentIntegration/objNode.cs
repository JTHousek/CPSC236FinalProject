using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAssessmentIntegration
{
    public class objNode
    {
        private string objective;
        private double students;
        private double maxScore;
        private double actualScore;
        private List<string> courseObjs = new List<string>();
        private objNode nextObjNode;

        public objNode()
        {
            objective = null;
            students = 0.0;
            maxScore = 0.1; //base set to 0.1 so division by zero does not occur
            actualScore = 0.0;
        }

        public objNode(string objective, double students, double maxScore, double actualScore)
        {
            this.objective = objective;
            this.students = students;
            this.maxScore = maxScore;
            this.actualScore = actualScore;
        }

        public void setObjective(string objective)
        {
            this.objective = objective;
        }

        public string getObjective()
        {
            return objective;
        }

        public void setStudents(double students)
        {
            this.students = students;
        }

        public double getStudents()
        {
            return students;
        }

        public void setMaxScore(double maxScore)
        {
            this.maxScore = maxScore;
        }

        public double getMaxScore()
        {
            return maxScore;
        }

        public void setActualScore(double actualScore)
        {
            this.actualScore = actualScore;
        }

        public double getActualScore()
        {
            return actualScore;
        }

        public void addCourseObj(string newCourseObj)
        {
            courseObjs.Add(newCourseObj);
        }

        public List<string> getCourseObj()
        {
            return courseObjs;
        }

        public void setNextObjNode(objNode nextObjNode)
        {
            this.nextObjNode = nextObjNode;
        }

        public objNode getNextObjNode()
        {
            return nextObjNode;
        }
    }
}
