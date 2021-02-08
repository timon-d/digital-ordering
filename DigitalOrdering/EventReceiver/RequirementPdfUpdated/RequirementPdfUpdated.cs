using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using InfoPathAttachmentEncoding;
using System.Xml.XPath;
using System.Xml;
using iTextSharp.text.pdf;
using System.IO;
using System.Net.Mail;
using System.Collections.Generic;

namespace DigitalOrdering.EventReceiver.RequirementPdfUpdated
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class RequirementPdfUpdated : SPItemEventReceiver
    {
        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);
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
                            //Festlegen von Variablen
                            //Initialisieren der "Helper"-Klasse. Diese enthält selbst geschriebene Methoden, die hier aufgerufen werden.
                            Helper helper = new Helper();
                            //Laden des aktuell geänderten bzw. unterschrieben Elements
                            SPListItem _currentItem = web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                            //Abspeichern von Spaltenwerten des Elements in Variablen; der String in den eckigen Klammern ist der Spaltenname
                            string requirementId = _currentItem["Title"].ToString();
                            string docId = _currentItem.Properties["_dlc_DocId"].ToString();
                            //Laden der Bibliothek "Temp" und des Unterordners, welcher als Namen die Auftragsnummer hat

                            //Angabe des Namens der Bibliothek der Auftragsformulare und deren Dateinamen
                            //string formLibrary = "BedarfsmeldungFormulare"; 
                            string formLibrary = "Bedarfsmeldung-Formulare";
                            string formFileName = requirementId + ".xml";
                            string xAttachment = "my:file";
                            string xReason = "/my:myFields/my:reason";
      
                            //Definieren einer vorerst leeren String-Variable
                            string prevSigned = "";
                            //Wenn die Spalte "SignaturePreviousSigned" nicht leer ist, wird der enthaltene Wert der Variable "davorsigniert" zugwiesen
                            if (_currentItem["SignaturePreviousSigned"] != null) prevSigned = _currentItem["SignaturePreviousSigned"].ToString();

                            //Auslesen der Empfängeradressen
                            string rcptRequirementFiles = "";
                            string spGroupRcptRequirementFiles = "Empfänger der Bedarfsmeldungen";
                            foreach (SPUser user in web.SiteGroups[spGroupRcptRequirementFiles].Users)
                            {
                                rcptRequirementFiles += user.Email.ToString() + ";";
                            }

                            string rcptBbc = "";
                            string spGroupRcptBbc = "E-Mails der Ereignisempfänger (BBC) ";
                            foreach (SPUser user in web.SiteGroups[spGroupRcptBbc].Users)
                            {
                                rcptBbc += user.Email.ToString() + ";";
                            }

                            //Laden des gerade signierten PDFs; zuerst wird es als SharePoint-Datei (SPFile) geladen, dann als Byte-Array.
                            SPFile requirementPdf = _currentItem.File;
                            byte[] requirementPdfByte = requirementPdf.OpenBinary();
                            //Initialisieren eines PDFReaders (von itextsharp) mit Verweis auf das eben geöffnete Byte-Array; der PDF-Reader kann PDF-Dateien nur lesen
                            using (PdfReader pdfreader = new PdfReader(requirementPdfByte))
                            {
                                //Laden der vorhandenen PDF-Formularfelder
                                AcroFields fields = pdfreader.AcroFields;
                                //Speichern der Namen aller signierten Felder in disem PDF in Liste "signatures". Die zu signierenden Felder heißen "Projektleiter" und "Verwaltung"; demnach kann dieser Wert leer, "Projektleiter" oder "ProjektleiterVerwaltung" sein
                                List<string> signatures = pdfreader.AcroFields.GetSignatureNames();
                                //Die gerade gespeicherten Feldnamen werden mit Komma getrennt (wenn mehrere vorhanden sind) und in eine neue String-Variable gespeichert. Somit wird aus "ProjektleiterVerwaltung" > "Projektleiter,Verwaltung"
                                string signed = string.Join(",", signatures.ToArray());
                                //Anhand dieser IF-Abfrage wird geprüft ob der Auftag überhaupt signiert wurde. Es wird der Spaltenwert ("SignaturePreviousSigned") mit den im PDF tatsächlich signierten Feldern verglichen
                                if (signed != prevSigned)
                                {
                                    //Starten des SharePoint-Designer-Workflows "Unterschriftenlauf" mit Parameter "signedfields". Mit diesem Parameter werden die Namen der Signierten Felder direkt als String mitgegeben.
                                    SPWorkflowAssociationCollection associationCollection = _currentItem.ParentList.WorkflowAssociations;
                                    foreach (SPWorkflowAssociation association in associationCollection)
                                    {
                                        if (association.Name == "Bedarfsmeldung - Genehmigungsvorgang")
                                        {
                                            //Überprüfen ob eine weitere Instanz des gleichen Workflows schon ausgeführt wird.
                                            //Das ist nötig, wenn eine Drittlandsbestellung gemacht wird. Dabei wird der Workflow solange angehalten/pausiert bis die Zollbeauftragten den Auftrag genehmigt haben.
                                            //Wenn vor der Genehmigung etwas am Auftrag geändert wird, muss der SP-Designer-WF erneut gestartet werden. Da die vorige Instanz des Workfows aber noch läuft und nur pausiert ist,
                                            //schlägt das Starten des Workflows fehl.
                                            foreach (SPWorkflow spworkflow in site.WorkflowManager.GetItemActiveWorkflows(_currentItem))
                                            {
                                                //Überprüft wird, ob die AssocationId übereinstimmt --> wenn ja, wird diese Instanz mit "CancelWorkflow" abgebrochen
                                                if (spworkflow.AssociationId == association.Id) SPWorkflowManager.CancelWorkflow(spworkflow);
                                            }
                                            association.AutoStartChange = true;
                                            association.AutoStartCreate = false;
                                            association.AssociationData = "<Data><Signiert>" + signed + "</Signiert></Data>";
                                            //web.Site.WorkflowManager.StartWorkflow(_currentItem, association, association.AssociationData);
                                            site.WorkflowManager.StartWorkflow(_currentItem, association, association.AssociationData);
                                        }
                                    }
                                    //Der SharePoint-Designer-Workflow geht nun alle vorhandenen Fälle durch und versendet die nötigen Benachrichtigungen und ändert das Status-Feld des Auftrags
                                    //Im Code muss nun nur noch der Fall beachtet werden, wenn Projektleiter und Verwaltung unterschrieben haben
                                    //In diesem Fall muss ein PDF für die Verwaltung und eins (nur wenn nötig) für den Empfang erstellt werden. Letzteres wird benötigt, um dem Empfang dieses per Mail zu versenden, damit es gefaxt oder per Post versendet werden kann. 
                                    //Das ist nur nötig, wenn die Bestellmethode NICHT "online" oder "Abholung" ist
                                    //Zu aller erst wird überprüft ob der beschriebene Fall eintritt: Die Variable "signiert" enthält alle Namen der Signaturfelder aus dem PDf, die unterschrieben sind. 
                                    if (signed == "Projektleiter")
                                    {
                                        SPList tempLibrary = web.Lists["Temp"];
                                        var tempFolder = web.GetFolder(String.Format("{0}/Temp/" + requirementId, web.Url));
                                        //Angabe des Namens der Bibliothek, in der die Begründungen gespeichert werden, und des Dateinamens der Begründungen
                                        string urlRequirementInTempFolder = tempLibrary.RootFolder.SubFolders[requirementId].Url + "/" + requirementId + ".pdf";

                                        string tempuploadurl = helper.TempOrdner(tempLibrary, tempFolder, requirementId);
                                        //Laden des Auftragsformulars
                                        byte[] xmlfile = helper.GetFileFromWeb(formLibrary, formFileName, web);
                                        Stream xmlmemstr = new MemoryStream(xmlfile);
                                        XmlDocument xmldoc = new XmlDocument();
                                        xmldoc.Load(xmlmemstr);
                                        XPathNavigator root = xmldoc.CreateNavigator();
                                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(new NameTable());
                                        //Angeben des NameSpaces: Die Werte werden direkt aus einem beispiel Formular entnommen (stehen am Anfang der XML-Datei)
                                        nsmgr.AddNamespace("my", "http://schemas.microsoft.com/office/infopath/2003/myXSD/2019-02-06T08:35:46");
                                        helper.UploadAttachments(root.Select("/my:myFields/my:attachments/my:attachment", nsmgr), xAttachment, nsmgr, docId, tempFolder, tempuploadurl, web);

                                        string reasonFolder = "Begruendungen";
                                        string reasonLibrary = "Begründungen";
                                        string reasonFileName = requirementId + "_Begruendung.pdf";
                                        string urlReasonInTempFolder = tempLibrary.RootFolder.SubFolders[requirementId].Url + "/" + reasonFileName;
                                        if (root.SelectSingleNode(xReason, nsmgr).ToString() != "")
                                        {
                                            byte[] reasonPdf = helper.GetFileFromWeb(reasonLibrary, reasonFileName, web);
                                            SPFile spFileReason = tempFolder.Files.Add(urlReasonInTempFolder, reasonPdf, true);
                                        }
                                        
                                        //Bedarfsmeldungs-Pdf
                                        byte[] flattenedRequirementPdf = helper.flattenPdfForm(requirementPdfByte);
                                        SPFile spFileRequirementPdf = tempFolder.Files.Add(urlRequirementInTempFolder, flattenedRequirementPdf, true);
                                        //Setzen des Spaltenwerts "Weitergabe" auf true, da der Lieferant den Auftragszettel immer erhalten muss. Wäre dieser WErt auf "false" würde Auftragszettel beim PDf für Fax oder Post fehlen. Der Wert ist in der SharePoint-Bibliothek "Temp" standardmäßig auf "false" gesetzt.
                                        //SPListItem spItemRequirementPdf = spFileRequirementPdf.Item;
                                        //spItemRequirementPdf.SystemUpdate();

                                        //diese IF-Abfrage ist nur reine Sicherheitsmaßnahme: Es wird geprüft ob im Temp-Ordner auch Dateien vorhanden sind. Da der Auftrag direkt davor in diesen Ordner kopiert wird sollte diese Bedingung immer zutreffen.
                                        if (tempFolder.ItemCount > 0)
                                        {
                                            SPListItemCollection itemCollection = helper.CollectFiles(tempFolder, tempLibrary, false);
                                            string mailText = "<div style='font-family: Frutiger LT COM 45 Light; font-size: 11pt;'>Guten Tag,<br><br>"
                                                + "die Bedarfsmeldung #" + requirementId + " wurde vom Projektleiter signiert. "
                                                + "Die Bedarfsmeldung und die zugehörigen Anlagen sind im Anhang enthalten.<br><br>"
                                                + "Name Besteller: " + helper.GetUserDisplayName(_currentItem, "Name Besteller") + "<br>"
                                                + "Projektleiter: " + helper.GetUserDisplayName(_currentItem, "Projektleiter") + "<br><br>"
                                                + "Dies ist eine automatische Benachrichtigung. Bitte antworten Sie nicht auf diese E-Mail.";
                                            helper.Mail("Bestellung", rcptRequirementFiles, rcptBbc, "Neue Bedarfsmeldung (#" + requirementId + ")", mailText, itemCollection, requirementId, site);
                                        }
                                        tempFolder.Delete();
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        public class Helper
        {
            public void Mail(string absendername, string empfänger, string bcc, string betreff, string text, SPListItemCollection itemCollection, string requirementId, SPSite site)
            {
                SmtpClient server = new SmtpClient(site.WebApplication.OutboundMailServiceInstance.Server.Address);
                MailMessage msg = new MailMessage();
                msg.From = new MailAddress(site.WebApplication.OutboundMailSenderAddress, absendername);
                msg.IsBodyHtml = true;
                msg.BodyEncoding = System.Text.Encoding.UTF8;
                string[] empfängerlist = empfänger.Split(';');
                foreach (string to in empfängerlist)
                {
                    if (to != "")
                    {
                        msg.To.Add(new MailAddress(to));
                    }
                }
                if (bcc != "")
                {
                    string[] bccList = bcc.Split(';');
                    foreach (string to in bccList)
                    {
                        if (to != "")
                        {
                            msg.Bcc.Add(new MailAddress(to));
                        }
                    }
                }
                msg.Subject = betreff;
                msg.Body = text;
                foreach (SPListItem li in itemCollection)
                {
                    byte[] ar = li.File.OpenBinary();
                    if (ar != null && ar.Length > 0)
                    {
                        MemoryStream ms = new MemoryStream(ar);
                        Attachment attachment = new Attachment(ms, li.File.Name.ToString());
                        msg.Attachments.Add(attachment);

                    }
                }
                server.Send(msg);
            }

            //Mit dieser Methode wird aus den Spalten "Name Besteller" und "Projektleiter" der "SPUser" erhalten; mit diesem kann auf Mail-Adresse und Anzeigename zugegriffen werden
            //Sie wird verwendet, um den Anzeigename in der Mail an die Zentrale verwenden zu können
            public string GetUserDisplayName(SPListItem listitem, string spalte)
            {
                //Laden des userFields aus dem übergebenem Element (listitem) und der Spalte, die einen Nutzer enthält (spalte)
                SPFieldUser userField = (SPFieldUser)listitem.Fields.GetField(spalte);
                SPFieldUserValue userFieldValue = (SPFieldUserValue)userField.GetFieldValue(listitem[spalte].ToString());
                SPUser user = userFieldValue.User;
                //Zuweisen des Anzeigenamens und Rückgabe desselben
                string anzeigename = user.Name;
                return anzeigename;
            }

            //Methode: Anlegen des Unterordners in der Bibliothek "Temp" (/Temp/<Auftragsnummer>).
            //Wenn der Unterordner bereits vorhanden ist (das tritt bei Formularänderungen auf), wird dieser gelösch und neu angelegt.
            //Außerdem wird die URL zum Unterordner als String zurückgegeben, die im Hauptprogramm dann verwendet wird, um Dateien in diesem abzuspeichern.
            public string TempOrdner(SPList tempbibliothek, SPFolder tempfolder, string auftragsnummer)
            {
                if (tempfolder.Exists)
                {
                    tempfolder.Delete();
                }
                var i = tempbibliothek.Items.Add("", SPFileSystemObjectType.Folder, auftragsnummer);
                //Ohne die Update()-Funktion wird kein Ordner angelegt!
                i.Update();
                //Anpassen des Ordnernamens
                tempfolder.Item["Title"] = auftragsnummer;
                tempfolder.Item.Update();
                string tempuploadurl = tempbibliothek.RootFolder.SubFolders[auftragsnummer].Url + "/";
                return tempuploadurl;
            }

            public byte[] GetFileFromWeb(string ordner, string dateiname, SPWeb web)
            {
                byte[] datei = new byte[0];
                SPList tempList = web.Lists[ordner];
                datei = web.GetFile(string.Format("{0}/{1}", tempList.RootFolder.Url, dateiname)).OpenBinary();
                //datei = web.Folders[ordner].Files[dateiname].OpenBinary();
                return datei;
            }

            public byte[] flattenPdfForm(byte[] byteArray)
            {
                byte[] result = new byte[0];
                using (MemoryStream ms = new MemoryStream())
                {
                    using (PdfReader pdfReader = new PdfReader(byteArray))
                    {
                        using (PdfStamper pdfStamper = new PdfStamper(pdfReader, ms))
                        {
                            pdfStamper.AcroFields.GenerateAppearances = true;
                            pdfStamper.AnnotationFlattening = true;
                            pdfStamper.FormFlattening = true;
                        }
                    }
                    result = ms.ToArray();
                }
                return result;
            }

            //Methode: Decodieren der Anlagen und Speichern in Bibliothek "Temp" im entsprechenden Unterordner
            public void UploadAttachments(XPathNodeIterator anlagen, string xanhang, XmlNamespaceManager nsmgr, string DocID, SPFolder tempfolder, string tempuploadurl, SPWeb web)
            {
                Helper helper = new Helper();
                //"Moven" durch jede Zeiler der wiederholten Tabelle der Anlagen mit MoveNext
                while (anlagen.MoveNext())
                {
                    //Verwenden von "try" und "catch", da der Code im "try"-Block fehlschlägt und das ganze Programm abbricht, wenn KEINE Anlagen vorhanden sind. Durch try - und catch wird nicht abgebrochen.
                    try
                    {
                        //Dekodieren des Anlagenfelds. Durch anlagen.current.selectsinglenode wird nicht die ganze Zeile, sondern nur das Feld in der Zeile zum Dekodieren ausgewählt. 
                        //Das ist wichtig, da noch das Boolean-Feld zur Weitergabe in jeder Zeile enthalten ist und das Dekodieren mit diesem Feld nicht möglich ist.
                        InfoPathAttachmentDecoder decoder = new InfoPathAttachmentDecoder(anlagen.Current.SelectSingleNode(xanhang, nsmgr).Value);
                        string fileNamewithoutextension = Path.GetFileNameWithoutExtension(decoder.Filename);
                        SPFile attachmentuploadfile = tempfolder.Files.Add(tempuploadurl + decoder.Filename, decoder.DecodedAttachment, true);
                        SPListItem attachmentitem = attachmentuploadfile.Item;
                        //Updaten des Elements damit Änderungen übernommen werden. SystemUpdate() ist eine "silent"-Änderung (dh. Änderungsdatum und Editor werden nicht erfasst). Wenn nur "Update()" verwendet wird, werden diese erfasst.
                        attachmentitem.SystemUpdate();
                    }
                    catch { }
                }
            }

            //Diese Methode sammelt Elemente aus dem Unterordner der SharePoint-Bibiothek "Temp"; diese "Sammlung" wird in einer anderen Methode verwendet um sie zu einer PDF-Datei zusammen zu fügen
            public SPListItemCollection CollectFiles(SPFolder ordner, SPList bibliothek, Boolean lieferant)
            {
                //Um Dokumente zu "sammeln" muss eine Abfrage(SPQuery) erstellt werden
                SPQuery Query = new SPQuery();
                //Dieser String ist wichtig, da nur durch ihn Elemente innerhalb eiens Ordners abgefragt werden können
                Query.ViewAttributes = "Scope=\"Recursive\"";
                string query = "<OrderBy><FieldRef Name=\"ID\" Ascending='False' /></OrderBy>";
                Query.Query = query;
                //Abzielen der Abfrage auf den im Aufruf mitgegeben Ordner
                Query.Folder = ordner;
                //Durchführen der Abfrage
                SPListItemCollection collListItems = bibliothek.GetItems(Query);
                //Rückgabe der "SPListItemCollection", damit dieser weiter verwendet werden kann.
                return collListItems;
            }
        }
    }
}