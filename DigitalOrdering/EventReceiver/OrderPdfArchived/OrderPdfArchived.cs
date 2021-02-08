using System;
using System.Linq;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Globalization;
using iTextSharp.text.pdf;
using DigitalOrdering.Classes;
using System.Text;
using System.IO;

namespace DigitalOrdering.EventReceiver.OrderPdfArchived
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class OrderPdfArchived : SPItemEventReceiver
    {
        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);

            //Ausführen des folgenden Codes nur, wenn ein PDF vorhanden ist
            if (properties.ListItem.Name.Contains(".pdf"))
            {
                //Durch folgende Zeile wird der enthaltene Code als SHAREPOINT\System ausgeführt und nicht unter dem Kontext des Benutzers, der den Auftrag unterschrieben hat
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    //Laden des aktuellen Webs; das ist nötig, um alle möglichen Operationen mit Bibliotheken und Elementen auszuführen
                    //dieses wird wieder geschlossen bzw. freigegeben, wenn die geschweifte Klammer von "using" geschlossen wird
                    using (SPSite site = new SPSite(properties.WebUrl))
                    {
                        using (SPWeb web = site.OpenWeb())
                        {
                            //Summe korrigieren
                            string orderPdfSumClm = "Summe";
                            string orderPdfSumField = "Summe";
                            string orderPdfNumberClm = "Auftragsnummer";

                            SPListItem _currentItem = web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                            string title = _currentItem.Title.ToString();
                            string logFile = @"C:\Bestellung\Logs\EventReceivers\OrderPdfArchived\" + title + ".log";
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine("OrderPdf " + title + " was added");
                            sb.AppendLine("Check if fields exist...");
                            bool sumFieldExists = _currentItem.Fields.ContainsField(orderPdfSumClm);
                            bool orderIdFieldExists = _currentItem.Fields.ContainsField(orderPdfNumberClm);
                            sb.AppendLine(orderPdfSumClm + ": " + sumFieldExists.ToString());
                            sb.AppendLine(orderPdfNumberClm + ": " + orderIdFieldExists.ToString());

                            if (sumFieldExists && orderIdFieldExists)
                            {
                                sb.AppendLine("Fields exist, go on...");
                                string orderId = _currentItem[orderPdfNumberClm].ToString();
                                decimal sumClmValue = decimal.Parse(_currentItem[orderPdfSumClm].ToString(), NumberStyles.Currency);
                                sb.AppendLine("Sum in collumn is: " + sumClmValue.ToString());

                                sb.AppendLine("Opening pdf...");
                                decimal sumPdfValue;
                                SPFile pdfFile = _currentItem.File;
                                byte[] pdfBytes = pdfFile.OpenBinary();
                                using (PdfReader pdfreader = new PdfReader(pdfBytes))
                                {
                                    sb.AppendLine("Pdf opened with iText...");
                                    AcroFields fields = pdfreader.AcroFields;
                                    sumPdfValue = decimal.Parse(fields.GetField(orderPdfSumField), NumberStyles.Currency);
                                    sb.AppendLine("Sum in pdf is: " + sumPdfValue.ToString());
                                }
                                sb.AppendLine("Checking if sum values differ...");
                                if (sumClmValue != sumPdfValue)
                                {
                                    sb.AppendLine("Sums are different. Assiging new sum (" + sumPdfValue + ") to collumn...");
                                    double sumPdfValueDouble = (double)sumPdfValue;
                                    _currentItem[orderPdfSumClm] = sumPdfValueDouble;
                                }
                                else
                                {
                                    sb.AppendLine("Sums are not different, carry on...");
                                }

                                //Gruppe eintragen
                                sb.AppendLine("Set the group collumn...");
                                string groupStr = "";
                                string orderPdfGroupClm = "Gruppe";
                                string orderFormGroupClm = "Gruppe";
                                string orderPdfIDDBClm = "IDDB";
                                string orderFormLibrary = "Auftragsformular";
                                string orderFormArchiveLibrary = "Auftragsformular-Archiv";

                                string orderFormIdStr = _currentItem[orderPdfIDDBClm].ToString();
                                sb.AppendLine("Got the ID of InfoPath form: " + orderFormIdStr);
                                string orderNumber = _currentItem[orderPdfNumberClm].ToString();
                                sb.AppendLine("Got the orderNumber");

                                decimal orderFormIdDec = Convert.ToDecimal(orderFormIdStr);
                                int orderFormId = Convert.ToInt32(orderFormIdDec);

                                SPList orderFormList = web.Lists.TryGetList(orderFormLibrary);
                                sb.AppendLine("Got the library " + orderFormList.Title.ToString());
                                
                                //Prüfen, ob das Auftragsformular noch nicht in das Archiv verschoben wurde
                                //Es wird die Anzahl aller Elemente, die die obrige ID haben, gespeichert.
                                sb.AppendLine("Trying to find Info-Path forms with ID " + orderFormId);
                                int countFoundOrderFormItems = (from SPListItem item in orderFormList.Items
                                                                where Convert.ToInt32(item["ID"]) == orderFormId
                                                                select item).Count();

                                if (countFoundOrderFormItems > 0)
                                {
                                    //Das Auftragsformular wurde noch nicht verschoben
                                    sb.AppendLine("Forms were found: " + countFoundOrderFormItems.ToString());
                                    sb.AppendLine("Trying to get the one form...");
                                    SPListItem orderFormItem = orderFormList.GetItemById(orderFormId);
                                    groupStr = orderFormItem[orderFormGroupClm].ToString();
                                    sb.AppendLine("Got form and group " + groupStr);
                                    SPGroup group = web.Groups[groupStr];
                                    string groupLogin = group.ID.ToString() + ";#" + group.LoginName.ToString();
                                    sb.AppendLine("Assigning group " + groupLogin + " to collumn");
                                    _currentItem[orderPdfGroupClm] = groupLogin;
                                }
                                else
                                //Das Auftragsformular wurde ins Archiv verschoben
                                {
                                    sb.AppendLine("No form was found...");
                                }
                                using (EventReceiverManager eventReceiverManager = new EventReceiverManager(true))
                                {
                                    sb.AppendLine("Try updating item " + _currentItem.Title.ToString());
                                    _currentItem.SystemUpdate(false);
                                    sb.AppendLine("Item updated");
                                }
                            }
                            else
                            {
                                sb.AppendLine("Required fields do not exist. Do nothing and quit.");
                            }
                            File.AppendAllText(logFile, sb.ToString());
                        }
                    }
                });
            }
        } 
    }
}