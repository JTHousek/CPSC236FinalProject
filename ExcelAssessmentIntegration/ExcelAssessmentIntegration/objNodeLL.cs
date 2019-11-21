using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAssessmentIntegration
{
    public class objNodeLL
    {
        private objNode head;
        private objNode tail;

        public objNode getHead()
        {
            return head;
        }

        public void setHead(objNode head)
        {
            this.head = head;
        }

        public objNode getTail()
        {
            return tail;
        }

        public void setTail(objNode tail)
        {
            this.tail = tail;
        }

    }

}
