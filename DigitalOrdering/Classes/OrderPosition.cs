using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    class OrderPosition
    {
        private static int positionCount = 0;
        private int positionNumber;
        private int quantity;
        private String productDescription;
        private double unitPrice;
        private double calculatedPrice;
        private String projectId;
        private Boolean isHazardousSubstance;

        public OrderPosition(int positionNumber, int quantity, String productDescription, double unitPrice, String projectId, Boolean isHazardousSubstance)
        {
            this.positionNumber = positionNumber;
            this.quantity = quantity;
            this.productDescription = productDescription;
            this.unitPrice = unitPrice;
            this.calculatedPrice = unitPrice * quantity;
            this.projectId = projectId;
            this.isHazardousSubstance = isHazardousSubstance;
            positionCount++;
        }

        public double GetCalculatedPrice()
        {
            return this.calculatedPrice;
        }

        public Boolean GetHazardousSubstance() 
        {
            return this.isHazardousSubstance;
        }

    }
}
