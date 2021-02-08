using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using DigitalOrdering.Classes;
using System.IO;
using System.Text;

namespace DigitalOrdering.EventReceiver.OrderFormArchived
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class OrderFormArchived : SPItemEventReceiver
    {
        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
            //Ausführen des folgenden Codes nur, wenn ein XML vorhanden ist
            if (properties.ListItem.Name.Contains(".xml"))
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
                            string orderNumberCol = "FileLeafRef";
                            string orderPdfLibraryName = "Auftragszettel-Archiv";
                            string attachmentClm = "Anlagen/Begründung";
                            string attachmentVIClm = "Anlagen/Begründung (VI)";
                            string attachmentDesc = "öffnen";

                            SPListItem _currentItem = properties.ListItem;
                            string title = _currentItem.Title.ToString();
                            string logFile = @"C:\Bestellung\Logs\EventReceivers\OrderPdfArchived\" + title + ".log";
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine("OrderForm " + title + " was added");
                            string orderNumber = Path.GetFileNameWithoutExtension(_currentItem[orderNumberCol].ToString());

                            SPList orderPdfLibrary = web.Lists.TryGetList(orderPdfLibraryName);
                            sb.AppendLine("Got list '" + orderPdfLibrary.Title + "'");
                            string folder = _currentItem.File.ParentFolder.ToString();
                            sb.AppendLine("Try getting folder '" + folder + "'...");
                            var nFolder = web.GetFolder(String.Format("{0}/Auftragsformular-Archiv/" + folder, web.Url));
                            sb.AppendLine("Got folder '" + nFolder.Url + "'");

                            sb.AppendLine("Querying for archived pdf... ");
                            SPQuery query = new SPQuery();
                            query.Query = "<Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>" + orderNumber + ".pdf</Value></Eq></Where>";
                            query.Folder = nFolder;
                            SPListItemCollection queriedOrderPdfsItems = orderPdfLibrary.GetItems(query);
                            int count = queriedOrderPdfsItems.Count;
                            sb.AppendLine("Found archived pdfs... " + count.ToString());
                            if (count == 1)
                            {
                                sb.AppendLine("Changing URLs for one pdf... ");
                                string attachmentUrl = SPUtility.ConcatUrls(web.ServerRelativeUrl, "/_layouts/FormServer.aspx?XMLLocation=/bestellung/" + folder + "/" + orderNumber + ".xml&OpenIn=Browser&DefaultView=Anlagen&Source=/bestellung/SitePages/Schliessen.aspx");
                                SPFieldUrlValue url = new SPFieldUrlValue();
                                url.Description = attachmentDesc;
                                url.Url = attachmentUrl;

                                string attachmentUrlVI = SPUtility.ConcatUrls(web.ServerRelativeUrl, "/_layouts/FormServer.aspx?XMLLocation=/bestellung/" + folder + "/" + orderNumber + ".xml&OpenIn=Browser&DefaultView=AnlagenWeitergabe&Source=/bestellung/SitePages/Schliessen.aspx");
                                SPFieldUrlValue urlVI = new SPFieldUrlValue();
                                urlVI.Description = attachmentDesc;
                                urlVI.Url = attachmentUrlVI;

                                sb.AppendLine("Url for collumn " + attachmentClm + " : " + url.Url.ToString());
                                sb.AppendLine("Url for collumn " + attachmentVIClm + " : " + urlVI.Url.ToString());

                                foreach (SPListItem orderPdf in queriedOrderPdfsItems)
                                {
                                    sb.AppendLine("Iterating through found pdfs, should be only one...");
                                    using (EventReceiverManager eventReceiverManager = new EventReceiverManager(true))
                                    {
                                        sb.AppendLine("Try updating item " + orderPdf.Title.ToString() + "...");
                                        orderPdf[attachmentClm] = url;
                                        orderPdf[attachmentVIClm] = urlVI;
                                        orderPdf.SystemUpdate(false);
                                        sb.AppendLine("Item updated");
                                    }
                                }
                            }
                            else
                            {
                                sb.AppendLine("There we're found more or less than one archived pdf... do nothing and quit");
                            }
                            File.AppendAllText(logFile, sb.ToString());
                        }
                    }
                });
            }
        }
    }
}