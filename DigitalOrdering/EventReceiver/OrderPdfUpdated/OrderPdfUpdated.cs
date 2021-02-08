using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using iTextSharp.text.pdf;
using iTextSharp.text;
using InfoPathAttachmentEncoding;
using System.Xml.XPath;
using System.Net.Mail;
using System.Xml;

namespace DigitalOrdering.EventReceiver.OrderPdfUpdated
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class OrderPdfUpdated : SPItemEventReceiver
    {
        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);
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
                            //Festlegen von Variablen
                            //Initialisieren der "Helper"-Klasse. Diese enthält selbst geschriebene Methoden, die hier aufgerufen werden.
                            Helper helper = new Helper();
                            //Laden des aktuell geänderten bzw. unterschrieben Elements
                            SPListItem _currentItem = web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                            //Abspeichern von Spaltenwerten des Elements in Variablen; der String in den eckigen Klammern ist der Spaltenname
                            string auftragsnummer = _currentItem["Auftragsnummer"].ToString();
                            string bestellmethode = _currentItem["Bestellmethode"].ToString();
                            SPFieldBoolean DrittlandBoolField = _currentItem.Fields["Drittland"] as SPFieldBoolean;
                            bool Drittland = (bool)DrittlandBoolField.GetFieldValue(_currentItem["Drittland"].ToString());
                            string DocID = _currentItem.Properties["_dlc_DocId"].ToString();
                            //Die Summe muss in ein double-Wert umgewandelt werden, damit er später richtig verwendet werden kann
                            double summe = Convert.ToDouble(_currentItem["Summe"]);
                            //Laden der Bibliothek "Temp" und des Unterordners, welcher als Namen die Auftragsnummer hat
                            SPList tempbibliothek = web.Lists["Temp"];
                            var tempfolder = web.GetFolder(String.Format("{0}/Temp/" + auftragsnummer, web.Url));
                            //Angabe des Namens der Bibliothek, in der die Begründungen gespeichert werden, und des Dateinamens der Begründungen
                            //string begruendungbibliothek = "Begruendungen";
                            string begruendungbibliothek = "Begründungen";
                            string begruendungdateiname = auftragsnummer + "_InterneBegruendung.pdf";
                            string auftragdesturl = tempbibliothek.RootFolder.SubFolders[auftragsnummer].Url + "/" + auftragsnummer + ".pdf";
                            string begruendungdesturl = tempbibliothek.RootFolder.SubFolders[auftragsnummer].Url + "/" + begruendungdateiname;
                            //Angabe des Namens der Bibliothek der Auftragsformulare und deren Dateinamen
                            string formularbibliothek = "Auftragsformular";
                            string formulardateiname = auftragsnummer + ".xml";
                            string xanhang = "my:Anhang";
                            string xweitergabe = "my:Weitergabe";

                            //Angeben des Pfades zum Ablegen des zusammengefassten PDFs für den Druck für die VW
                            //Bei Pfadangaben mit Backslashes muss entweder vor dem Anführungszeichen ein @ geschrieben werden ODER doppelte Backslashes verwendet werden (z.B. "C:\\ipm-int\\Bestellungen\\")
                            string pdfpath = @"C:\Bestellung\PrintTemp\PDF\";
                            //Angeben des Druckers über den gedruckt wird; hier ist einfach der Freigabename anzugeben
                            //string printer = @"\\ipm-int\Bestellungen\";
                            string printer = @"\\printserver\A-3-31-SW-Buero\";
                            //Definieren einer vorerst leeren String-Variable
                            string davorsigniert = "";
                            //Wenn die Spalte "SignaturePreviousSigned" nicht leer ist, wird der enthaltene Wert der Variable "davorsigniert" zugwiesen
                            if (_currentItem["SignaturePreviousSigned"] != null) davorsigniert = _currentItem["SignaturePreviousSigned"].ToString();


                            //Auslesen der Empfänger-Adressen aus den SharePoint-Gruppen
                            string rcptFaxPostalOrders = "";
                            string spGroupRcptFaxPostalFiles = "Empfänger der Fax - und Postbestellungen";
                            foreach (SPUser user in web.SiteGroups[spGroupRcptFaxPostalFiles].Users)
                            {
                                rcptFaxPostalOrders += user.Email.ToString() + ";";
                            }
                            Console.WriteLine(rcptFaxPostalOrders);

                            //Die Drittlands-Mail inkl. Anlagen soll seit 13.02.2019 laut Verwaltung nur noch an das Einkaufs-Postfach gehen.
                            string rcptThirdCountryOrders = "";
                            string spGroupRcptCustomFiles = "Empfänger der Drittlandsbestellungen";
                            foreach (SPUser user in web.SiteGroups[spGroupRcptCustomFiles].Users)
                            {
                                rcptThirdCountryOrders += user.Email.ToString() + ";";
                            }

                            string rcptBbc = "";
                            string spGroupRcptBbc = "E-Mails der Ereignisempfänger (BBC) ";
                            foreach (SPUser user in web.SiteGroups[spGroupRcptBbc].Users)
                            {
                                rcptBbc += user.Email.ToString() + ";";
                            }

                            //Erstellen der byte-Arrays für die PDF-Dateien die durch Mergen/Zusammenfügen erstellt werden; eins für den Empfang zum faxen, das andere für den Ausdruck für die Verwaltung
                            byte[] mergeresultbyte = new byte[0];
                            byte[] mergeresultbytevw = new byte[0];
                            //Festlegen des Dateinamens unter dem das PDF für den Ausdurck für die Verwaltung im lokalen Dateisystem abgespeichert wird.
                            string mergefilenamevw = auftragsnummer + "_vw.pdf";

                            //Laden des gerade signierten PDFs; zuerst wird es als SharePoint-Datei (SPFile) geladen, dann als Byte-Array.
                            SPFile auftragszettel = _currentItem.File;
                            byte[] auftragszettelbyte = auftragszettel.OpenBinary();
                            //Initialisieren eines PDFReaders (von itextsharp) mit Verweis auf das eben geöffnete Byte-Array; der PDF-Reader kann PDF-Dateien nur lesen
                            using (PdfReader pdfreader = new PdfReader(auftragszettelbyte))
                            {
                                //Laden der vorhandenen PDF-Formularfelder
                                AcroFields fields = pdfreader.AcroFields;
                                //Speichern der Namen aller signierten Felder in disem PDF in Liste "signatures". Die zu signierenden Felder heißen "Projektleiter" und "Verwaltung"; demnach kann dieser Wert leer, "Projektleiter" oder "ProjektleiterVerwaltung" sein
                                List<string> signatures = pdfreader.AcroFields.GetSignatureNames();
                                //Die gerade gespeicherten Feldnamen werden mit Komma getrennt (wenn mehrere vorhanden sind) und in eine neue String-Variable gespeichert. Somit wird aus "ProjektleiterVerwaltung" > "Projektleiter,Verwaltung"
                                string signiert = string.Join(",", signatures.ToArray());
                                //Anhand dieser IF-Abfrage wird geprüft ob der Auftag überhaupt signiert wurde. Es wird der Spaltenwert ("SignaturePreviousSigned") mit den im PDF tatsächlich signierten Feldern verglichen
                                if (signiert != davorsigniert)
                                {
                                    //Starten des SharePoint-Designer-Workflows "Unterschriftenlauf" mit Parameter "signedfields". Mit diesem Parameter werden die Namen der Signierten Felder direkt als String mitgegeben.
                                    SPWorkflowAssociationCollection associationCollection = _currentItem.ParentList.WorkflowAssociations;
                                    foreach (SPWorkflowAssociation association in associationCollection)
                                    {
                                        if (association.Name == "Unterschriftenlauf")
                                        {
                                            association.AutoStartChange = true;
                                            association.AutoStartCreate = false;
                                            association.AssociationData = "<Data><Signiert>" + signiert + "</Signiert></Data>";
                                            //web.Site.WorkflowManager.StartWorkflow(_currentItem, association, association.AssociationData);
                                            site.WorkflowManager.StartWorkflow(_currentItem, association, association.AssociationData);
                                        }
                                    }
                                    //Der SharePoint-Designer-Workflow geht nun alle vorhandenen Fälle durch und versendet die nötigen Benachrichtigungen und ändert das Status-Feld des Auftrags
                                    //Im Code muss nun nur noch der Fall beachtet werden, wenn Projektleiter und Verwaltung unterschrieben haben
                                    //In diesem Fall muss ein PDF für die Verwaltung und eins (nur wenn nötig) für den Empfang erstellt werden. Letzteres wird benötigt, um dem Empfang dieses per Mail zu versenden, damit es gefaxt oder per Post versendet werden kann. 
                                    //Das ist nur nötig, wenn die Bestellmethode NICHT "online" oder "Abholung" ist
                                    //Zu aller erst wird überprüft ob der beschriebene Fall eintritt: Die Variable "signiert" enthält alle Namen der Signaturfelder aus dem PDf, die unterschrieben sind. 
                                    if ((signiert == "Projektleiter,Verwaltung") || (signiert == "Verwaltung,Projektleiter"))
                                    {
                                        string tempuploadurl = helper.TempOrdner(tempbibliothek, tempfolder, auftragsnummer);
                                        //Laden des Auftragsformulars
                                        byte[] xmlfile = helper.GetFileFromWeb(formularbibliothek, formulardateiname, web);
                                        Stream xmlmemstr = new MemoryStream(xmlfile);
                                        XmlDocument xmldoc = new XmlDocument();
                                        xmldoc.Load(xmlmemstr);
                                        XPathNavigator root = xmldoc.CreateNavigator();
                                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(new NameTable());
                                        //Angeben des NameSpaces: Die Werte werden direkt aus einem beispiel Formular entnommen (stehen am Anfang der XML-Datei)
                                        nsmgr.AddNamespace("my", "http://schemas.microsoft.com/office/infopath/2003/myXSD/2016-04-22T15:49:19");
                                        helper.AnlagenHochladen(root.Select("/my:meineFelder/my:Anlagen/my:Anlage", nsmgr), xanhang, xweitergabe, nsmgr, DocID, tempfolder, tempuploadurl, web);

                                        //Hochladen der Begründung und des aktuellen Auftragzettels in den Temp-Ordner
                                        //Begründung
                                        if (root.SelectSingleNode("/my:meineFelder/my:Begruendung", nsmgr).ToString() != "")
                                        {
                                            byte[] fattenedBegruendung = helper.flattenPdfForm(helper.GetFileFromWeb(begruendungbibliothek, begruendungdateiname, web));
                                            SPFile copybegruendung = tempfolder.Files.Add(begruendungdesturl, fattenedBegruendung, true);
                                            SPListItem copybegruendungitem = copybegruendung.Item;
                                            copybegruendungitem["Auftragsnummer"] = auftragsnummer;
                                            //Übernehmen der Änderung. Durch "SystemUpdate" wird diese Änderung nicht in SharePoint selbst dokumentiert (in Spalten "geändert von" und "geändert").
                                            copybegruendungitem.SystemUpdate();
                                        }

                                        //Auftragszettel
                                        byte[] flattendAuftrag = helper.flattenPdfForm(auftragszettelbyte);
                                        SPFile copyauftrag = tempfolder.Files.Add(auftragdesturl, flattendAuftrag, true);
                                        //Setzen des Spaltenwerts "Weitergabe" auf true, da der Lieferant den Auftragszettel immer erhalten muss. Wäre dieser WErt auf "false" würde Auftragszettel beim PDf für Fax oder Post fehlen. Der Wert ist in der SharePoint-Bibliothek "Temp" standardmäßig auf "false" gesetzt.
                                        SPListItem copyauftragitem = copyauftrag.Item;
                                        copyauftragitem["Weitergabe"] = true;
                                        copyauftragitem.SystemUpdate();

                                        //diese IF-Abfrage ist nur reine Sicherheitsmaßnahme: Es wird geprüft ob im Temp-Ordner auch Dateien vorhanden sind. Da der Auftrag direkt davor in diesen Ordner kopiert wird sollte diese Bedingung immer zutreffen.
                                        if (tempfolder.ItemCount > 0)
                                        {
                                            //Zusammenführen aller PDF-Dokumente für die Verwaltung
                                            //SPListItemCollection ist eine Sammlung von SharePoint-Elementen. Welche Elemente ausgewählt werden, ist in der Helper-Methode (ganz unten) nach zu lesen. Mit dem Übergabewert "false" wird bewirkt, dass nicht beachtet wird, welche Anlagen zur Weitergabe an den Lieferanten gewählt sind - es wird alles einfach gedruckt.
                                            SPListItemCollection pdfcollection = helper.CollectPDF(tempfolder, tempbibliothek, false);
                                            //Die zuvor erstellte Sammlung von Elementen ("pdfcollection") wird nun in einer weiteren Helper-Methode verwendet, um die PDf-Dateien zu kombinieren. Mit dem Übergabewert "false" wird bewirkt, dass die AEB's NICHT mit ausgedruckt werden.
                                            mergeresultbytevw = helper.Merge(pdfcollection, false, auftragsnummer);

                                            //Zusammenführen der PDf-Dokumente für den Lieferanten
                                            //Gemerged wird nur, wenn über Fax oder Post bestellt wird, damit die Zentrale eine Datei hat, die sie versenden/ausdrucken können. Wenn "online" oder per "Abholung" bestellt wird, oder wenn es sich um ein Drittland handelt, findet dieser Vorgang nicht statt.
                                            if ((bestellmethode.Contains("online") == false) && (bestellmethode.Contains("Abholung") == false) && Drittland == false)
                                            {
                                                //Sammeln der Elemente; diesmal mit dem Übergabewert "true", damit nur Anlagen berücksichtigt werden, die auch zur Weitergabe für den Lieferanten ausgewählt sind
                                                pdfcollection = helper.CollectPDF(tempfolder, tempbibliothek, true);
                                                //Zusammenführen der gerade gesammelten Elemente; diesmal mit dem Übergabewert "true", damit die AEBs erhalten bleiben
                                                mergeresultbyte = helper.Merge(pdfcollection, true, auftragsnummer);
                                                //Kopieren des PDF in einen lokalen Ordner auf dem SharePoint-Server
                                                File.WriteAllBytes(pdfpath + auftragsnummer + "_Kopie.pdf", mergeresultbyte);
                                                string faxnummer = "<br>";
                                                if ((bestellmethode == "Fax") && (_currentItem["Fax"] != null))
                                                {
                                                    faxnummer = "Fax: " + _currentItem["Fax"].ToString() + "<br><br>";
                                                }
                                                string zentraletext = "<div style='font-family: Frutiger LT COM 45 Light; font-size: 11pt;'>Guten Tag,<br><br>"
                                                + "der Auftrag #" + auftragsnummer + " wurde vom Projektleiter und der Verwaltung signiert. Als Bestellmethode wurde " + bestellmethode + " gewählt. "
                                                + "Der Auftrag ist im Anhang enthalten. Hierbei handelt es sich nur um eine Kopie für den Lieferanten.<br><br>"
                                                + "Name Besteller: " + helper.Anzeigename(_currentItem, "Profil:NameBesteller") + "<br>"
                                                + "Projektleiter: " + helper.Anzeigename(_currentItem, "Profil:Projektleiter") + "<br>"
                                                + "Firma: " + _currentItem["Firma"].ToString() + "<br>"
                                                + faxnummer
                                                + "Dies ist eine automatische Benachrichtigung. Bitte antworten Sie nicht auf diese E-Mail.";
                                                helper.Mail("Bestellung", rcptFaxPostalOrders, rcptBbc, "Auftrag #" + auftragsnummer + " - Signaturvorgang abgeschlossen", zentraletext, mergeresultbyte, auftragsnummer, site);
                                            }
                                        }
                                        else
                                        {
                                            mergeresultbytevw = auftragszettelbyte; //Auch Dieser Else-Block ist nur reine Sicherheitsmaßnahme.
                                        }
                                        //Kopieren des PDFs für VW in einen lokalen Ordner (siehe Variable "printpath") auf dem SharePoint-Server. 
                                        File.WriteAllBytes(pdfpath + mergefilenamevw, mergeresultbytevw);
                                        //Hochladen des PDFs für VW in Ordner "Print" in Bibliothek "Temp"
                                        //desturl = tempbibliothek.RootFolder.SubFolders["Merged"].Url + "/" + mergefilenamevw;
                                        //SPFile uploadmerged = tempfolder.Files.Add(desturl, File.ReadAllBytes(pdfpath + mergefilenamevw), true);
                                        //Schicken der PDF-Datei für VW an den Drucker (siehe Variable "printer")
                                        if (Drittland == false)
                                        {
                                            File.Copy(pdfpath + mergefilenamevw, printer + mergefilenamevw);
                                        }
                                        else
                                        {
                                            string drittlandtext = "<div style='font-family: Frutiger LT COM 45 Light; font-size: 11pt;'>Guten Tag,<br><br>"
                                                + "der Signaturvorang für den Auftrag (#" + auftragsnummer + ") wurde abgeschlossen. Da es sich um eine Bestellung aus dem <b>Drittland</b> handelt, wird der Auftrag <b>ohne automatischen Ausdruck</b> an den Einkauf weitergeleitet. "
                                                + "Der Auftrag ist im Anhang enthalten.<br><br>"
                                                + "Name Besteller: " + helper.Anzeigename(_currentItem, "Profil:NameBesteller") + "<br>"
                                                + "Projektleiter: " + helper.Anzeigename(_currentItem, "Profil:Projektleiter") + "<br>"
                                                + "Firma: " + _currentItem["Firma"].ToString() + "<br>"
                                                + "Bestellmethode: " + bestellmethode + "<br><br>"
                                                + "Dies ist eine automatische Benachrichtigung. Bitte antworten Sie nicht auf diese E-Mail.";
                                            helper.Mail("Bestellung", rcptThirdCountryOrders, rcptBbc, "Drittland: Auftrag #" + auftragsnummer + " - Signaturvorgang abgeschlossen", drittlandtext, mergeresultbytevw, auftragsnummer, site);
                                        }
                                        //Löschen des PDFs für VW aus lokalem Ordner - NICHT MEHR NÖTIG, da der Temp-Ordner wöchentlich aufgeräumt wird. Deßhalb können beide Dateien aus Troubleshooting-Gründen aufgehoben werden.
                                        //File.Delete(pdfpath + mergefilenamevw);
                                        //Löschen des temporären Ordners in Bibliothek "Temp"
                                        tempfolder.Delete();
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
    }
    //Alle verwendeten Methoden sind hier definiert

    public class Helper
    {
        public void Mail(string absendername, string empfänger, string bcc, string betreff, string text, byte[] anlage, string auftragsnummer, SPSite site)
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
            if (anlage != null && anlage.Length > 0)
            {
                //Hinzufügen der zusammengeführen PDf-Datei als Anhang: Initilaisieren eines MemoryStreams der mit "mergeresultbyte" das zusammgengeführte PDf für den Lieferanten verwendet.
                MemoryStream attms = new MemoryStream(anlage);
                //Name des Anhangs <auftragsnummer + "_Kopie.pdf"
                Attachment attachment = new Attachment(attms, auftragsnummer + "_Kopie.pdf");
                //eigentliches hinzufügen des Anhangs
                msg.Attachments.Add(attachment);
            }
            server.Send(msg);
        }

        //Mit dieser Methode wird aus den Spalten "Name Besteller" und "Projektleiter" der "SPUser" erhalten; mit diesem kann auf Mail-Adresse und Anzeigename zugegriffen werden
        //Sie wird verwendet, um den Anzeigename in der Mail an die Zentrale verwenden zu können
        public string Anzeigename(SPListItem listitem, string spalte)
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
        public void AnlagenHochladen(XPathNodeIterator anlagen, string xanhang, string xweitergabe, XmlNamespaceManager nsmgr, string DocID, SPFolder tempfolder, string tempuploadurl, SPWeb web)
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
                    Boolean weitergabe = anlagen.Current.SelectSingleNode(xweitergabe, nsmgr).ValueAsBoolean;
                    Boolean portfolio = helper.PDFPortfolio(decoder.DecodedAttachment, fileNamewithoutextension, weitergabe, tempfolder, tempuploadurl, DocID);
                    if (portfolio != true)
                    {
                        //Hochladen in /Temp/<Auftragsnummer>
                        SPFile attachmentuploadfile = tempfolder.Files.Add(tempuploadurl + DocID + "_" + decoder.Filename, decoder.DecodedAttachment, true);
                        SPListItem attachmentitem = attachmentuploadfile.Item;
                        attachmentitem["Weitergabe"] = weitergabe;
                        //Updaten des Elements damit Änderungen übernommen werden. SystemUpdate() ist eine "silent"-Änderung (dh. Änderungsdatum und Editor werden nicht erfasst). Wenn nur "Update()" verwendet wird, werden diese erfasst.
                        attachmentitem.SystemUpdate();
                    }
                }
                catch { }
            }
        }

        public Boolean PDFPortfolio(byte[] pdf, string filename, Boolean weitergabe, SPFolder tempfolder, string tempuploadurl, string DocID)
        {
            Boolean portfolio;
            using (PdfReader reader = new PdfReader(pdf))
            {
                PdfReader.unethicalreading = true;
                PdfDictionary documentNames = null;
                PdfDictionary embeddedFiles = null;
                PdfDictionary fileArray = null;
                PdfDictionary file = null;
                PRStream stream = null;
                PdfDictionary catalog = reader.Catalog;
                documentNames = (PdfDictionary)PdfReader.GetPdfObject(catalog.Get(PdfName.NAMES));
                if (documentNames != null)
                {
                    //Erster Check: Portfolio ist vorhanden
                    embeddedFiles = (PdfDictionary)PdfReader.GetPdfObject(documentNames.Get(PdfName.EMBEDDEDFILES));
                    if (embeddedFiles != null)
                    {
                        //Zweiter Check: Portfolio ist wirklich vorhanden
                        portfolio = true;
                        PdfArray filespecs = embeddedFiles.GetAsArray(PdfName.NAMES);
                        for (int i = 0; i < filespecs.Size; i++)
                        {
                            i++;
                            fileArray = filespecs.GetAsDict(i);
                            file = fileArray.GetAsDict(PdfName.EF);
                            int filecount = 0;
                            foreach (PdfName key in file.Keys)
                            {
                                stream = (PRStream)PdfReader.GetPdfObject(file.GetAsIndirectObject(key));
                                //string attachedFileName = folderName + fileArray.GetAsString(key).ToString();
                                filecount++;
                                string attachedFileName = DocID + "_" + filename + "_ExportFromPortfolio_" + i + ".pdf";
                                byte[] attachedFileBytes = PdfReader.GetStreamBytes(stream);
                                SPFile attachmentuploadfile = tempfolder.Files.Add(tempuploadurl + attachedFileName, attachedFileBytes, true);
                                SPListItem attachmentitem = attachmentuploadfile.Item;
                                attachmentitem["Weitergabe"] = weitergabe;
                                attachmentitem.SystemUpdate();
                            }
                        }
                    }
                    else
                    {
                        //Zweiter Check: Doch kein Portfolio vorhanden
                        portfolio = false;
                    }
                }
                else
                {
                    //Erster check: Kein Portfolio vorhanden
                    portfolio = false;
                }
            }
            return portfolio;
        }

        //Diese Methode sammelt Elemente aus dem Unterordner der SharePoint-Bibiothek "Temp"; diese "Sammlung" wird in einer anderen Methode verwendet um sie zu einer PDF-Datei zusammen zu fügen
        public SPListItemCollection CollectPDF(SPFolder ordner, SPList bibliothek, Boolean lieferant)
        {
            //Um Dokumente zu "sammeln" muss eine Abfrage(SPQuery) erstellt werden
            SPQuery Query = new SPQuery();
            //Dieser String ist wichtig, da nur durch ihn Elemente innerhalb eiens Ordners abgefragt werden können
            Query.ViewAttributes = "Scope=\"Recursive\"";
            //Definieren von zwei Abfragen. Eine für den Ausdruck für die Verwaltung (queryvw) und eine für den Lieferanten (querylieferant)
            //In beiden Fällen wird sortiert nach der ID (absteigend). Die ID wird von SharePoint vergeben sobald eine Datei hochgeladen wird. Dadurch ist der Auftragszettel, der eben kopiert wurde, immer an erster Stelle, da er die höchste ID hat. 
            //Die Begründung wird ebenfalls immer nach Anlagen hochgeladen. Dadurch ist sie an zweiter Stelle, was auch so gewünscht ist.
            //Es wird nach allen Dateien die im Dateityp "pdf" enthalten haben abgefragt
            string queryvw = "<OrderBy><FieldRef Name=\"ID\" Ascending='False' /></OrderBy><Where><Contains><FieldRef Name='File_x0020_Type'/><Value Type='text'>pdf</Value></Contains></Where>";
            //Bei der Abfrage für den Lieferanten wird zusätzlich nach der Spalte "Weitergabe" geschaut und ob diese auf "true/wahr" steht; nur diese Elemente werden berücksichtigt
            string querylieferant = "<OrderBy><FieldRef Name=\"ID\" Ascending='False' /></OrderBy><Where><And><Contains><FieldRef Name='File_x0020_Type'/><Value Type='text'>pdf</Value></Contains><Contains><FieldRef Name='Weitergabe'/><Value Type='Boolean'>1</Value></Contains></And></Where>";
            //Standardmäßig wird die Abfrage für die Verwaltung ausgewählt...
            Query.Query = queryvw;
            //...außer der Methode wird der Boolean-Wert "true" mitgegeben; in diesem Fall wird die Abfrage für den Lieferanten ausgewählt
            if (lieferant == true) Query.Query = querylieferant;
            //Abzielen der Abfrage auf den im Aufruf mitgegeben Ordner
            Query.Folder = ordner;
            //Durchführen der Abfrage
            SPListItemCollection collListItems = bibliothek.GetItems(Query);
            //Rückgabe der "SPListItemCollection", damit dieser weiter verwendet werden kann.
            return collListItems;
        }

        //Methode zum Zusammenführen/Mergen der gesammelten PDFs; Zusammengeführt wird, was in "SPListItemCollection" mitgegeben wird (vorherige Methode); Rückgabewert ist das Ergebnis als byteArray
        public byte[] Merge(SPListItemCollection listItems, Boolean eab, string auftragsnummer)
        {
            //Initialiseren eines neuen Byte-Arrays; diesem wird später das erstellte PDf als MemoryStream übergeben.
            byte[] result = new byte[0];
            //"Öffnen" eines neuen Memorystreams; dieser bleibt solange offen, bis die geschweifte Klammer von "using" wieder geschlossen wird
            using (MemoryStream ms = new MemoryStream())
            {
                //Initialisieren eines neuen itextsharp-Dokuments; auch dieses bleibt solange geöffnet, bis die geschweifte Klammer von using wieder geschlossen wird
                using (Document document = new Document())
                {
                    //Starten von PdfCopy  (von itextsharp), es werden das zuvor erstellte Dokument und der geöffnete MemoryStream zugwiesen.
                    using (PdfCopy copy = new PdfCopy(document, ms))
                    {
                        document.Open();
                        //Mit eienr Foreach-Schleife wird durch alle SharePoint-Elemente innerhalb der mitgegeben "SPListItemCollection" durchgegangen
                        foreach (SPListItem li in listItems)
                        {
                            //Durch initialisieren eines PDF-Readers (itextsharp) wird jedes einzelne PDf zum Lesen geöffnet
                            using (PdfReader pdfreader = new PdfReader(li.File.OpenBinaryStream()))
                            {
                                //Dieser Befehl ist wichtig, da mit ihm passwortgeschützte PDFs, oder PDFs, die das Zusammenführen nicht erlauben, TROTZDEM zusammengeführt werden können
                                PdfReader.unethicalreading = true;
                                //Abspeichern der Seitenanzahl des gerade geöffneten PDFs
                                int n = pdfreader.NumberOfPages;
                                //Wenn die aktuelle Datei der Auftragszettel selbst ist UND wenn beim Aufruf der Methode für "eab" "false" mitgebgen wird, wird nur die erste Seite hinzugefügt.
                                //Somit bleiben die AEBs raus. Das wird für die ERstellung des PDfs für den Lieferanten verwendet. Beim Erstellen des Dokuments für die Verwaltung wird eab auf "true" gesetzt.
                                if ((li.File.Name == auftragsnummer + ".pdf") && (eab == false))
                                {
                                    //Befehl zum importieren, der ersten Seite des Auftragszettels
                                    copy.AddPage(copy.GetImportedPage(pdfreader, 1));
                                }
                                //Wenn es sich um jede andere Datei handelt UND eab == true ist, werden immer alle Seiten importiert.
                                else
                                {
                                    //jede Seite wird importiert, der Zähler wird mit jeder Seite hochgezählt, bis keine Seite mehr vorhanden ist.
                                    for (int page = 0; page < n; )
                                    {
                                        copy.AddPage(copy.GetImportedPage(pdfreader, ++page));
                                    }
                                }
                            }
                        }
                    }
                }
                //Speichern des MemoryStreams in das ByteArray.
                result = ms.ToArray();
            }
            //Rückgabe des erstellten PDFs als ByteArray zur Weiterverwendung
            return result;
        }
    }
}
