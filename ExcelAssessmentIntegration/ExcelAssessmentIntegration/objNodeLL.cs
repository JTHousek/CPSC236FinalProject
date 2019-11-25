//James Houseknecht, Cain Kaltenbaugh, Ethan Mooney
//jth1012, cxk1047, eam1020
//ExcelAssessmentIntegration
//objNodeLL.cs
//Start Date: October 1, 2019
//End Date: December 1, 2019

namespace ExcelAssessmentIntegration
{
    public class objNodeLL
    {
        private objNode head;   //head of the objective node linked list
        private objNode tail;   //tail of the objective node linked list

        //constructor for the objective node linked list
        public objNodeLL()
        {
            head = null;
            tail = null;
        }

        //clear all the values in the current linked list
        public void clearObjLL()
        {
            objNode currNode = head; //used to iterate through the linked list'
            while (currNode != null)
            {
                currNode.clearAverage();
                currNode.clearStudents();
                currNode.setTotalStudents(0);
                currNode.setWeightedAverage(0.0);
                currNode = currNode.getNextObjNode();
            }

        }

        //returns the head of the list
        public objNode getHead()
        {
            return head;
        }

        //sets the head of the list to a node
        public void setHead(objNode head)
        {
            this.head = head;
        }

        //returns the tail of the list
        public objNode getTail()
        {
            return tail;
        }

        //sets the tail of the list to a node
        public void setTail(objNode tail)
        {
            this.tail = tail;
        }

    }

}
