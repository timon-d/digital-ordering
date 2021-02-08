using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    class DigitalOrder
    {
        //Bestellübergreifende Variablen:
        private static String[] customsAgents;
        private static String[] orderApprovers;
        private static String[] hazardousSubstancesCommissioned;
        private static String[] laserProtectionOfficers;
        private static String[] centralOffice;

        //Variablen innerhalb einer Bestellung
        private int orderId;
        private Boolean isUVGO;
        private Boolean isInvest;
        private String company;
        private String customerNumber;
        private Boolean isThirdCountry;
        private String projectId;
        private Boolean isMultipleProjects;
        private int projectsBookingType = 0; //0 = nur ein Project; 1 = pro Position; 2 = prozentual
        private String customer;
        private String customerPhone;
        private String customerPhoneExtension;
        private String projectLeader;
        private String department;
        private Boolean isLaser;
        private Boolean isHazardousSubstance;

        private String quotationNumber;
        private OrderAttachment[] attachments;
        private int orderingMethod; //0 = online/E-Mail; 1 = Fax; 2 = Post; 3 = Abholung

        private OrderPosition[] positions;

        private double totalPrice;
        private String reason;
        private String noteToSupplier;
        private String noteToProcurement;
        private DateTime wishDeliveryDate;

        private DateTime creationDate;
        private DateTime modificationDate;
        private Boolean isCustomsAppoved;
        private Boolean isFirstSignature;
        private Boolean isSecondSignature;
        private Boolean isOrderApproved;
        private DateTime firstSignatureDate;
        private DateTime secondSignatureDate;
        private DateTime approvalDate;

        private String sigmaTransactionNumber;
        
        public DigitalOrder()
        {

        }

        public void GetTotalPrice() 
        {
            double result = 0;
            foreach(OrderPosition position in positions)
            {
                result += position.GetCalculatedPrice();
            }
            this.totalPrice = result;
        }

        public double GetTotalPrice() 
        {
            return this.totalPrice;
        }
    }
}
