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
