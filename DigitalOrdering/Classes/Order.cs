using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Collections;
using System.Globalization;
using System.Threading;
using InfoPathAttachmentEncoding;
using System.Net.Mail;
using System.Drawing.Drawing2D;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace DigitalOrdering.Classes.Order
{
    public class Order
    {
        //SharePoint
        public SPSite site;
        public SPWeb web;
        public String libraryNameOrders = "Auftragszettel";
        public String libraryNameTemp = "Temp";
        public SPFolder tempFolder;
        public String libraryNameReasons = "Begründungen";
        public String folderNameReasons = "Begruendungen";
        public SPList libraryOrders;
        public SPList libraryTemp;
        public SPList libraryReasons;
        public SPListItem listItemForm;
        public SPListItem listItemOrder;
        public SPListItem listItemReason;
        public SPFile orderFile;
        public SPFile reasonFile;

        public String templateFolder = "Vorlage";
        public String templateGerman = "Vorlage.pdf";
        public String templateEnglish = "Vorlage_engl.pdf";
        public String templateThirdCountry = "Vorlage_Zoll.pdf";
        public String templateHS8 = "Vorlage_HS8.pdf";
        public String templateThirdCountryHS8 = "Vorlage_Zoll_HS8.pdf";
        public String templateReason = "Vorlage_Begruendung.pdf";
        public String checkInComment = "Check-In durch das System";
        public String pdfFileNameOrder;
        public String pdfFileNameReason;
        public String collumnOrderOrderNumber = "Auftragsnummer";
        public String collumnOrderOrderNr = "Auftrags-Nr";
        public String collumnOrderDate = "Datum";
        public String collumnOrderCompany = "Firma";
        public String collumnOrderLaser = "Laser";
        public String collumnOrderDanger = "Gefahrstoff";
        public String collumnOrderIDDB = "IDDB";
        public String collumnOrderName = "Name Besteller";
        public String collumnOrderProfileName = "Profil:NameBesteller";
        public String collumnOrderProjectLeader = "Projektleiter";
        public String collumnOrderProfileProjectLeader = "Profil:Projektleiter";
        public String collumnOrderProject = "Projekt-Nr.";
        public String collumnOrderTotal = "Summe";
        public String collumnOrderOrderingMethod = "Bestellmethode";
        public String collumnOrderReason = "Begründung";
        public String collumnOrderAttachments = "Anlagen/Begründung";
        public String collumnOrderAttachmentsVI = "Anlagen/Begründung (VI)";
        public String collumnOrderFax = "Fax";
        public String collumnOrderGroup = "Gruppe";
        public String collumnOrderSignDate = "Datum Zweitunterschrift";
        public String collumnOrderApproval = "ZollGenehmigung";
        public String collumnOrderThirdCountry = "Drittland";

        public String collumnReasonOrderNumber = "Auftragsnummer";
        public String collumnTempTitle = "Title";
        public String collumnTempToSupplier = "Weitergabe";
        public String collumnFormOrderNumber = "Auftrags-Nr";
        public String collumnFormName = "Name Besteller";
        public String collumnFormProjectleader = "Projektleiter";
        public String workflow1Name = "Daten übergeben";
        public String workflow2Name = "Unterschriftenlauf";

        //Infopath-Formular / XML
        public XPathNavigator root;
        public XmlNamespaceManager nsmgr;
        public String nameSpacePrefix = "my";
        public String nameSpaceUri = "http://schemas.microsoft.com/office/infopath/2003/myXSD/2016-04-22T15:49:19";
        public String xOrderNumber = "/my:meineFelder/@my:Auftrags-Nr";
        public String xCompany = "/my:meineFelder/my:Firma";
        public String xThirdCountry = "/my:meineFelder/my:Drittland";
        public String xGKA301 = "/my:meineFelder/my:GKA301";
        public String xLanguageEnglish = "/my:meineFelder/my:Englisch";
        public String xOthers = "/my:meineFelder/my:Sonstiges";
        public String xDanger = "/my:meineFelder/my:Gefahrstoff";
        public String xLaser = "/my:meineFelder/my:Laser";
        public String xStreet = "/my:meineFelder/my:PostBereich/my:Straße";
        public String xZipcode = "/my:meineFelder/my:PostBereich/my:PLZ";
        public String xLocation = "/my:meineFelder/my:PostBereich/my:Ort";
        public String xCountry = "/my:meineFelder/my:PostBereich/my:Land";
        public String xOrderingMethod = "/my:meineFelder/my:Bestellmethode";
        public String xFax = "/my:meineFelder/my:FaxBereich/my:Fax";
        public String xTelephone = "/my:meineFelder/my:TelefonGesamt";
        public String xOfferNumber = "/my:meineFelder/my:Angebotsnummer";
        public String xCustomerNumber = "/my:meineFelder/my:Kundennummer";
        public String xTotal = "/my:meineFelder/my:Summe";
        public String xReason = "/my:meineFelder/my:Begruendung";
        public String xAnnotation = "/my:meineFelder/my:Anmerkung";
        public String xProjectNumber = "/my:meineFelder/my:Projekt-Nr";
        public String xMultipleProjects = "/my:meineFelder/my:MehrProjekte";
        public String xMultipleProjectsType = "/my:meineFelder/my:MehrProjekteArt";
        public String xAttachments = "/my:meineFelder/my:Anlagen/my:Anlage";
        public XPathNodeIterator xIteratorAttachments;
        public String xAttachment = "my:Anhang";
        public String xAttachmentsToSupplier = "my:Weitergabe";
        public String xToSupplierChange = "/my:meineFelder/my:WeitergabeÄnderung";
        public String xPositions = "/my:meineFelder/my:Positionen/my:Position";
        public XPathNodeIterator xIteratorPositions;
        public String xPositionsNumber = "my:Positionsnummer";
        public String xPositionsAmount = "my:Bestellmenge";
        public String xPositionsCustomsTariffsNumber = "my:ZollBereich/my:Warentarifnummer";
        public String xPositionsProjectNumber = "my:ProjektNrBereich/my:Projekt-Nr.Position";
        public String xPositionsText = "my:Text";
        public String xPositionsUnitPrice = "my:Einzelpreis";
        public String xPositionsTotalPrice = "my:Gesamtpreis";
        public String xProjects = "/my:meineFelder/my:ProjekteProzentualBereich/my:ProjekteProzentual/my:Projekt";
        public XPathNodeIterator xIteratorProjects;
        public String xProjectsPercent = "my:Prozent";
        public String xProjectsNumber = "my:Projekt-Nr.Prozent";
        public String xAttachmentsCheck = "/my:meineFelder/my:AnlagenCheck";
        public String xGroup = "/my:meineFelder/my:Gruppe";

        //PDF-Einstellungen
        public String fieldOrderOrderNumber = "Auftrags-Nr";
        public String fieldOrderCompany = "Firma";
        public String fieldOrderOthers = "Sonstiges";
        public String fieldOrderOfferNumber = "Angebotsnummer";
        public String fieldOrderTelephone = "Telefon";
        public String fieldOrderCustomerNumber = "Kundennummer";
        public String fieldOrderName = "Name Besteller";
        public String fieldOrderDate = "Datum";
        public String fieldOrderTotal = "Summe";
        public String fieldOrderAnnotation = "Anmerkung";
        public String fieldOrderProjectNumber = "Projekt-Nr";
        public String fieldOrderPointerProjectNumber = "PointerPN";
        public String fieldOrderSignatureProjectleader = "Projektleiter";
        public String fieldOrderSignatureAdministration = "Verwaltung";
        public String fieldReasonOrderNumber = "AuftragBegruendung";
        public String fieldReasonReason = "Begruendung";

        public Font frutigerFont = new Font(BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.CP1252, BaseFont.EMBEDDED), 10);
        public int currencyFormat = 1031;

        public String orderNumber;
        public Boolean thirdCountry;
        public Boolean thirdCountryApproval = false;
        public Boolean gka301;
        public Boolean hs8;
        public Boolean languageEnglish = false;
        public Boolean reasonEmpty = true;
        public Boolean toSupplierChange;
        public Boolean attachmentsCheck;
        public String danger = "Nein";
        public String laser = "Nein";
        public byte[] orderTemplate = new byte[0];
        public byte[] reasonTemplate = new byte[0];
        public byte[] orderPdf = new byte[0];
        public byte[] reasonPdf = new byte[0];
        public String tempUploadUrl;

        public void LoadInfoPathForm()
        {
            byte[] xmlFile = listItemForm.File.OpenBinary();
            Stream xmlMemoryStream = new MemoryStream(xmlFile);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlMemoryStream);
            root = xmlDoc.CreateNavigator();
            nsmgr = new XmlNamespaceManager(new NameTable());
            nsmgr.AddNamespace(nameSpacePrefix, nameSpaceUri);
            orderNumber = root.SelectSingleNode(xOrderNumber, nsmgr).ToString();
            pdfFileNameOrder = orderNumber + ".pdf";
            pdfFileNameReason = orderNumber + "_InterneBegruendung.pdf";
            gka301 = root.SelectSingleNode(xGKA301, nsmgr).ValueAsBoolean;
            hs8 = !gka301;
            thirdCountry = root.SelectSingleNode(xThirdCountry, nsmgr).ValueAsBoolean;
            if (!thirdCountry) thirdCountryApproval = true;
            if (root.SelectSingleNode(xDanger, nsmgr).ToString() != "false") danger = "Ja";
            if (root.SelectSingleNode(xLaser, nsmgr).ToString() != "false") laser = "Ja";
            attachmentsCheck = root.SelectSingleNode(xAttachmentsCheck, nsmgr).ValueAsBoolean;
            if (root.SelectSingleNode(xReason, nsmgr).ToString() != "") reasonEmpty = false;
            toSupplierChange = root.SelectSingleNode(xToSupplierChange, nsmgr).ValueAsBoolean;
            if (root.SelectSingleNode(xLanguageEnglish, nsmgr) != null) languageEnglish = root.SelectSingleNode(xLanguageEnglish, nsmgr).ValueAsBoolean;
        }

        //Methode: Zum Konvertieren der Währungen aus dem Auftragsformular vom Datentyp "double" in einen String, um mit diesem das PDF-Formular zu füllen.
        public String CurrencyToString(double doubleSource)
        {
            string result = null;
            CultureInfo ci = new CultureInfo(currencyFormat);
            result = String.Format(ci, "{0:C}", doubleSource);
            return result;
        }

        //Methode: Anlegen des Unterordners in der Bibliothek "Temp" (/Temp/<Auftragsnummer>).
        //Wenn der Unterordner bereits vorhanden ist (das tritt bei Formularänderungen auf), wird dieser gelösch und neu angelegt.
        //Außerdem wird die URL zum Unterordner als String zurückgegeben, die im Hauptprogramm dann verwendet wird, um Dateien in diesem abzuspeichern.
        public void CreateTempFolder()
        {
            libraryTemp = web.Lists[libraryNameTemp];
            tempFolder = web.GetFolder(String.Format("{0}/Temp/" + orderNumber, web.Url));
            if (tempFolder.Exists)
            {
                tempFolder.Delete();
            }
            var i = libraryTemp.Items.Add("", SPFileSystemObjectType.Folder, orderNumber);
            //Ohne die Update()-Funktion wird kein Ordner angelegt!
            i.Update();
            tempFolder.Item[collumnTempTitle] = orderNumber;
            tempFolder.Item.Update();
            tempUploadUrl = libraryTemp.RootFolder.SubFolders[orderNumber].Url + "/";
        }

        //Methode: Es wird eine der PDF-Vorlagen aus der Bibliothek "Vorlage" je nach angegebenen Werten im Auftragsformular geladen und als Byte-Array zurück gegeben.
        //Mit web.folders[<Bibliothek>].Files[<Dateiname>].OpenBinary() wird eine Datei binär bzw. als Byte-Array geöffnet
        public void GetPdfTemplates()
        {
            if (thirdCountry == true && hs8 == true)
            {
                orderTemplate = web.Folders[templateFolder].Files[templateThirdCountryHS8].OpenBinary();
            }
            else if (thirdCountry == true && hs8 == false)
            {
                orderTemplate = web.Folders[templateFolder].Files[templateThirdCountry].OpenBinary();
            }
            else if (thirdCountry == false && hs8 == true)
            {
                orderTemplate = web.Folders[templateFolder].Files[templateHS8].OpenBinary();
            }
            else if (thirdCountry == false && hs8 == false)
            {
                orderTemplate = web.Folders[templateFolder].Files[templateGerman].OpenBinary();
            }
        }

        //Methode: Diese Methode befüllt eine Zelle mit Werten und fügt diese einer Tabelle hinzu
        public void AddPDFCell(string text, int textAlign, int border, int paddingLeft, int paddingRight, float minimumHeight, PdfPTable table)
        {
            //Erstellen einer neuen Zelle
            PdfPCell cell = new PdfPCell(new Phrase(text, frutigerFont));
            //Definieren der Parameter der Zelle. Die genauen Werte werden beim Methoden-Aufruf übergeben
            cell.Border = border;
            cell.PaddingLeft = paddingLeft;
            cell.PaddingRight = paddingRight;
            cell.MinimumHeight = minimumHeight;
            cell.HorizontalAlignment = textAlign;
            //Durch AddCell wird die Zelle tatsächlich einer Tabelle hinzugefügt
            table.AddCell(cell);
        }

        //Methode: Befüllen des PDF-Formlars mit Werten
        public void CreateOrderPdf()
        {
            //initialiseren eines MemoryStreams, welcher dem PDF-Stamper (siehe nächster Befehl) übergeben wird
            using (Document document = new Document(PageSize.A4))
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    //Folgende drei Initalisierungen werden in Kombination mit using(){} gemacht. Somit gelten die Initialisierungen nur innerhalb des using-Abschnitts.
                    //Das ist eine C#-Best-Practice, da ansonsten bspw. das bearbeitete PDF für immer in einem "Die Datei wird schon verwendet"-Zustand bleibt und nicht verwendet werden kann.
                    //initialisieren des PDF-Readers von itextsharp (zum Lesen von PDFs); diesem wird die Vorlage des Auftragszettels als ByteArray übergeben
                    using (PdfReader pdfReader = new PdfReader(orderTemplate))
                    {
                        //initialisieren des PDF-Stampers von itextsharp (zum Bearbeiten von PDFs)
                        using (PdfStamper pdfStamper = new PdfStamper(pdfReader, memoryStream, '\0', true))
                        {
                            //In folgendem Abschnitt werden die vorhandenen PDF-Formularfelder mit den Werten des Formuars befüllt
                            //Syntax: fields.SetFied(<Feldname im PDF>, <String>);
                            //Als zu übergebender String wird mit "root.SelectSingleNode(<XPath zum Feld im InfoPath-Formular>).ToString() der Wert eines Formularfeldes innerhalb von InfoPath übergeben.
                            //Da manche Werte (bsp. Auftragsnummer, Firma...) auch schon als Variablen oder innerhalb von Listenspalten vorhanden sind, werden diese verwendet.
                            //Zugriff auf Spaltenwert: _currentItem[<Spaltenname>].ToString();
                            //Mit pdfstamper.AcroFields werden alle Formularfelder des PDFs geladen
                            AcroFields fields = pdfStamper.AcroFields;
                            fields.SetField(fieldOrderCompany, root.SelectSingleNode(xCompany, nsmgr).ToString());
                            //Der Block unterhalb des Firmennamens auf dem Auftragszettel bietet Platz für Ansprechpartner, Adresse, PLZ, Ort, Bestellmethode und Faxnummer.
                            //Immer wenn ein Wert vorhanden ist, wird ein Zeilenumbruch mit "\n" hinzugefügt. Wenn nicht soll nicht umgebrochen werden, damit keien unnötigen Textlücken entehen.
                            //Alle einelnen InfoPath-Felder werden in den String "infobox" gespeichert und dieser wird in das PDF-Formularfeld "Sonstiges" geschrieben.
                            string infobox = "";
                            if (root.SelectSingleNode(xOthers, nsmgr).ToString() != "") infobox = root.SelectSingleNode(xOthers, nsmgr).ToString() + "\n";
                            if (root.SelectSingleNode(xStreet, nsmgr).ToString() != "") infobox += root.SelectSingleNode(xStreet, nsmgr).ToString() + "\n";
                            if (root.SelectSingleNode(xZipcode, nsmgr).ToString() != "") infobox += root.SelectSingleNode(xZipcode, nsmgr).ToString();
                            if (root.SelectSingleNode(xLocation, nsmgr).ToString() != "") infobox += " " + root.SelectSingleNode(xLocation, nsmgr).ToString() + "\n";
                            if (root.SelectSingleNode(xCountry, nsmgr).ToString() != "") infobox += root.SelectSingleNode(xCountry, nsmgr).ToString();
                            if (root.SelectSingleNode(xOrderingMethod, nsmgr).ToString() == "Fax") infobox += "\nFax: " + root.SelectSingleNode(xFax, nsmgr).ToString();
                            else if ((root.SelectSingleNode(xOrderingMethod, nsmgr).ToString().Contains("online")) || (root.SelectSingleNode(xOrderingMethod, nsmgr).ToString().Contains("Abholung"))) infobox += "\n" + root.SelectSingleNode(xOrderingMethod, nsmgr).ToString();
                            fields.SetField(fieldOrderOthers, infobox);
                            fields.SetField(fieldOrderOrderNumber, orderNumber);
                            fields.SetField(fieldOrderOfferNumber, root.SelectSingleNode(xOfferNumber, nsmgr).ToString());
                            fields.SetField(fieldOrderTelephone, root.SelectSingleNode(xTelephone, nsmgr).ToString());
                            fields.SetField(fieldOrderCustomerNumber, root.SelectSingleNode(xCustomerNumber, nsmgr).ToString());
                            fields.SetField(fieldOrderName, listItemForm[collumnFormName].ToString());
                            fields.SetField(fieldOrderDate, System.DateTime.Now.ToShortDateString());
                            fields.SetField(fieldOrderTotal, CurrencyToString(root.SelectSingleNode(xTotal, nsmgr).ValueAsDouble));
                            if (hs8 == false) fields.SetField(fieldOrderAnnotation, root.SelectSingleNode(xAnnotation, nsmgr).ToString());
                            fields.SetField(fieldOrderProjectNumber, root.SelectSingleNode(xProjectNumber, nsmgr).ToString());

                            //Prüfen ob auf mehrere Projekte gebucht wird: wenn ja wird das Formularfeld "Projekt-Nr" unten links auf dem Auftragszettel mit leerem String gefüllt
                            if (root.SelectSingleNode(xMultipleProjects, nsmgr).ToString() == "true")
                            {
                                fields.SetField(fieldOrderPointerProjectNumber, "");
                            }

                            //Alle Felder im PDF (außer Signaturfelder) schreibgeschützt machen
                            //Foreach-Schleife geht alle Formularfelder innerhalb des PDFs durch
                            foreach (var field in fields.Fields)
                            {
                                //field.Key entspricht dem Namen des Formularfelds im PDF; "Projektleiter" und "Verwaltung" sind die Namen der Signaturfelder; die sollen nicht schreibgeschützt sein
                                if ((field.Key != fieldOrderSignatureProjectleader) && (field.Key != fieldOrderSignatureAdministration))
                                {
                                    //Befehl zum Setzen des Felds auf "ReadOnly"
                                    fields.SetFieldProperty(field.Key.ToString(), "setfflags", PdfFormField.FF_READ_ONLY, null);
                                }
                            }

                            //Prüfen, ob im Falle von mehreren Projekten, diese prozentual angegeben wurden (dh. 20% auf Projekt xxxxxx und 89% auf Projekt xxxxxx)
                            if (root.SelectSingleNode(xMultipleProjectsType, nsmgr).Value == "prozentual")
                            {
                                //Tabelle mit zwei Spalten
                                PdfPTable pdfTableProjects = new PdfPTable(2);
                                pdfTableProjects.SetTotalWidth(new float[] { 30, 50 });
                                //Eintragen von Überschriften für beide Spalten
                                AddPDFCell("%", 1, 3, 0, 0, 10f, pdfTableProjects);
                                AddPDFCell("Projekt-Nr.", 1, 3, 0, 0, 10f, pdfTableProjects);
                                //In dieser while-Schleife wird die wiederholte Tabelle Zeile für Zeile durchgegangen und jeder Wert mit der helper-Methode "PdfZelle" der Tabelle hinzugefügt
                                xIteratorProjects = root.Select(xProjects, nsmgr);
                                while (xIteratorProjects.MoveNext())
                                {
                                    AddPDFCell(xIteratorProjects.Current.SelectSingleNode(xProjectsPercent, nsmgr).Value, 1, 3, 0, 0, 10f, pdfTableProjects);
                                    AddPDFCell(xIteratorProjects.Current.SelectSingleNode(xProjectsNumber, nsmgr).Value, 1, 3, 0, 0, 10f, pdfTableProjects);
                                }
                                //Durch "WriteSelectedRows" wird die Tabelle erst in das PDF geschrieben; ohne den Befehl existiert keine Tabelle.
                                pdfTableProjects.WriteSelectedRows(0, 2, 0, xIteratorProjects.Count + 1, 37, 120, pdfStamper.GetOverContent(1));
                            }
                            //Die Positionen werden in Form einer Tabelle auf den Auftragszettel geschrieben
                            //initialisieren einer neuen Tabelle; die Zahl in Klammern gibt die Spaltenanzahl an
                            PdfPTable pdfTablePositions = new PdfPTable(5);
                            //Angeben der Breite der jeweiligen Spalten; Spalte 1 ist 34px breit, Spalte 2 ist 45px breit usw.
                            pdfTablePositions.SetTotalWidth(new float[] { 34, 45, 294, 85, 86 });
                            //Keine Borders in der Tabelle, die Tabellenrahmen sind schon im PDf vorhanden
                            pdfTablePositions.DefaultCell.Border = Rectangle.NO_BORDER;
                            //"positionszeilen" ist der "XPath" zur Wiederholten Tabelle innerhalb des InfoPath-Formulars(wurde wie die anderen Variablen am Anfang des Hauptprogramms definiert)
                            //über "MoveNext" wird jede Positionsreihe durchgegangen
                            xIteratorPositions = root.Select(xPositions, nsmgr);
                            while (xIteratorPositions.MoveNext())
                            {
                                //Die einzelnen Zellen werden für die momentan gewählte Position mit der selbst angelegeten helper-Methode "PdfZelle" mit Werten befüllt
                                AddPDFCell(xIteratorPositions.Current.SelectSingleNode(xPositionsNumber, nsmgr).Value, 1, 0, 4, 0, 20f, pdfTablePositions);
                                AddPDFCell(xIteratorPositions.Current.SelectSingleNode(xPositionsAmount, nsmgr).Value, 1, 0, 4, 0, 20f, pdfTablePositions);
                                string sonstiges = "";
                                //Falls eine Warentarifnummer für diese Position angegeben ist, wird diese in Variable "sonstiges" gespeichert
                                if (xIteratorPositions.Current.SelectSingleNode(xPositionsCustomsTariffsNumber, nsmgr).Value != String.Empty)
                                {
                                    //Mit "\n" wird ein Zeilenumbruch generiert
                                    sonstiges = "\nHarmonized Tariff Number: " + xIteratorPositions.Current.SelectSingleNode(xPositionsCustomsTariffsNumber, nsmgr).Value;
                                }
                                //Falls eine Projektnummer für diese Position angegeben ist, wird diese in Variable "sonstiges" gespeichert (bzw. zur Warentarifnummer hinzugefügt)
                                if (xIteratorPositions.Current.SelectSingleNode(xPositionsProjectNumber, nsmgr).Value != String.Empty)
                                {
                                    //Mit "\n" wird ein Zeilenumbruch generiert
                                    sonstiges = sonstiges + "\nProjekt-Nr.: " + xIteratorPositions.Current.SelectSingleNode(xPositionsProjectNumber, nsmgr).Value;
                                }
                                //Zur Zelle der Spalte "Text" wird der angegebene Text im Formular PLUS die Variabel "sonstiges" eingetragen. In dieser ist, falls vorhanden Warentarif - und Projektnummer.
                                AddPDFCell(xIteratorPositions.Current.SelectSingleNode(xPositionsText, nsmgr).Value + sonstiges, 0, 0, 5, 0, 20f, pdfTablePositions);
                                //Einzelpreis und Gesamtpreis müssen mit der helper-Methode "ConvertCurrency" in einen String umgewandelt werden.
                                //Mit "PdfPCell.ALIGN_RIGHT" wird alles rechts eingerückt
                                AddPDFCell(CurrencyToString(xIteratorPositions.Current.SelectSingleNode(xPositionsUnitPrice, nsmgr).ValueAsDouble), PdfPCell.ALIGN_RIGHT, 0, 0, 5, 20f, pdfTablePositions);
                                AddPDFCell(CurrencyToString(xIteratorPositions.Current.SelectSingleNode(xPositionsTotalPrice, nsmgr).ValueAsDouble), PdfPCell.ALIGN_RIGHT, 0, 0, 5, 20f, pdfTablePositions);
                            }
                            //Durch "WriteSelectedRows" wird die Tabelle erst in das PDF geschrieben; ohne den Befehl existiert keine Tabelle.
                            pdfTablePositions.WriteSelectedRows(0, 5, 0, xIteratorPositions.Count, 37, 560, pdfStamper.GetOverContent(1));
                        }
                    }
                    //Bis zu dieser STelle sind alle Editierungen des PDFS nur in einem MemoryStream vorhanden. Durch "ToArray()" wird dieser MemoryStream in ein ByteArray umgewandelt.
                    orderPdf = memoryStream.ToArray();
                }
            }
        }

        public void UploadFile(String targetLibraryName, String fileName, byte[] file, Boolean overwrite)
        {
            SPList targetLibrary = web.Lists[targetLibraryName];
            String destUrl = SPUtility.ConcatUrls(web.Url, targetLibrary.RootFolder.Url) + "/" + fileName;
            SPFile newFile = web.GetFile(destUrl);
            if (newFile.Exists)
            {
                if (newFile.CheckOutType != SPFile.SPCheckOutType.None)
                {
                    newFile.CheckIn(checkInComment);
                }
                using (EventReceiverManager eventReceiverManager = new EventReceiverManager(true))
                {
                    newFile.Delete();
                }
            }
            SPFieldUserValue userValue = new SPFieldUserValue(web, listItemForm[SPBuiltInFieldId.Author].ToString());
            SPUser author = userValue.User;
            SPUserToken userToken = author.UserToken;
            using (SPSite impSite = new SPSite(web.Site.ID, userToken))
            {
                using (SPWeb impWeb = impSite.OpenWeb())
                {
                    SPList impTargetLibrary = impWeb.Lists[targetLibraryName];
                    String impDestUrl = SPUtility.ConcatUrls(impWeb.Url, impTargetLibrary.RootFolder.Url) + "/" + fileName;
                    SPFile upload = impTargetLibrary.RootFolder.Files.Add(impDestUrl, file, overwrite);
                }
            }
        }

        public void UpdateListItemOrder()
        {
            orderFile = web.Folders[libraryNameOrders].Files[pdfFileNameOrder];
            listItemOrder = orderFile.Item;
            listItemOrder[collumnOrderOrderNumber] = orderNumber;   //String
            listItemOrder[collumnOrderOrderNr] = orderNumber;   //String
            listItemOrder[collumnOrderIDDB] = listItemForm.ID.ToString();   //String
            listItemOrder[collumnOrderDanger] = danger; //String
            listItemOrder[collumnOrderLaser] = laser;  //String
            listItemOrder[collumnOrderCompany] = root.SelectSingleNode(xCompany, nsmgr).Value.ToString(); //String
            listItemOrder[collumnOrderProject] = root.SelectSingleNode(xProjectNumber, nsmgr).Value.ToString(); //String
            listItemOrder[collumnOrderDate] = listItemForm["Erstellt"]; //Datum
            listItemOrder[collumnOrderSignDate] = DateTime.Now;
            listItemOrder[collumnOrderFax] = root.SelectSingleNode(xFax, nsmgr).Value.ToString(); //String
            listItemOrder[collumnOrderGroup] = web.SiteGroups[root.SelectSingleNode(xGroup, nsmgr).Value.ToString()];  //Personenfeld
            listItemOrder[collumnOrderOrderingMethod] = root.SelectSingleNode(xOrderingMethod, nsmgr).Value.ToString(); //String            
            listItemOrder[collumnOrderApproval] = thirdCountryApproval; //Boolean
            listItemOrder[collumnOrderThirdCountry] = thirdCountry;   //Boolean
            if (!reasonEmpty || attachmentsCheck)
            {
                SetHyperlinkField(collumnOrderAttachments, "öffnen", "Anlagen");
                SetHyperlinkField(collumnOrderAttachmentsVI, "öffnen", "AnlagenWeitergabe");
            }
            else
            {
                listItemOrder[collumnOrderAttachments] = null;
                listItemOrder[collumnOrderAttachmentsVI] = null;
            }
            listItemOrder.SystemUpdate();
        }

        public SPUser GetUserProfileByDisplayName(String displayName)
        {
            SPPrincipalInfo pinfo = SPUtility.ResolvePrincipal(web, displayName, SPPrincipalType.User, SPPrincipalSource.All, web.Users, false);
            SPUser user = web.Users[pinfo.LoginName];
            return user;
        }

        public void SetHyperlinkField(String collumn, String description, String defaultView)
        {
            var urlFieldValue = new SPFieldUrlValue();
            urlFieldValue.Description = description;
            urlFieldValue.Url = listItemForm.ParentList.ParentWebUrl.ToString() + "/_layouts/FormServer.aspx?XMLLocation=" + listItemForm.ParentList.ParentWebUrl.ToString() + "/" + listItemForm.Url.ToString() + "&OpenIn=Browser&DefaultView=" + defaultView + "&Source=" + listItemForm.ParentList.ParentWebUrl.ToString() + "/SitePages/Schliessen.aspx";
            listItemOrder[collumn] = urlFieldValue;
        }

        //Methode: Decodieren der Anlagen und Speichern in Bibliothek "Temp" im entsprechenden Unterordner
        public void UploadAttachments()
        {
            xIteratorAttachments = root.Select(xAttachments, nsmgr);
            //"Moven" durch jede Zeiler der wiederholten Tabelle der Anlagen mit MoveNext
            while (xIteratorAttachments.MoveNext())
            {
                //Verwenden von "try" und "catch", da der Code im "try"-Block fehlschlägt und das ganze Programm abbricht, wenn KEINE Anlagen vorhanden sind. Durch try - und catch wird nicht abgebrochen.
                try
                {
                    //Dekodieren des Anlagenfelds. Durch anlagen.current.selectsinglenode wird nicht die ganze Zeile, sondern nur das Feld in der Zeile zum Dekodieren ausgewählt. 
                    //Das ist wichtig, da noch das Boolean-Feld zur Weitergabe in jeder Zeile enthalten ist und das Dekodieren mit diesem Feld nicht möglich ist.
                    InfoPathAttachmentDecoder decoder = new InfoPathAttachmentDecoder(xIteratorAttachments.Current.SelectSingleNode(xAttachment, nsmgr).Value);
                    //Hochladen in /Temp/<Auftragsnummer>
                    SPFile attachmentUploadFile = tempFolder.Files.Add(tempUploadUrl + decoder.Filename, decoder.DecodedAttachment, true);
                }
                catch { }
            }
        }

        //Methode: Zum Erstellen der internen Begründung als PDF
        public void CreateReasonPdf()
        {
            //Laden der PDF-Vorlage in ein ByteArray. "Vorlagenordner" gibt die Bibliothek an und "vorlage" gibt den zu Namen der zu ladenden Datei an.
            reasonPdf = web.Folders[templateFolder].Files[templateReason].OpenBinary();
            //Initialisieren eines PDF-Readers von itextsharp (zum Lesen)
            using (PdfReader pdfReader = new PdfReader(reasonPdf))
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (PdfStamper pdfStamper = new PdfStamper(pdfReader, memoryStream, '\0', true))
                    {
                        //Mit pdfstamper.AcroFields werden alle Formularfelder des PDFs geladen
                        AcroFields formFields = pdfStamper.AcroFields;
                        //Eintragen der Auftragsnummer und der Begründung in entsprechende Formularfelder (Syntax: fields.SetField(<Name des PDF-Feldes>, <string>);)
                        formFields.SetField(fieldReasonOrderNumber, orderNumber);
                        formFields.SetField(fieldReasonReason, root.SelectSingleNode(xReason, nsmgr).Value.ToString());
                        //Stellen der Felder auf Schreibgeschützt
                        formFields.SetFieldProperty(fieldReasonOrderNumber, "setfflags", PdfFormField.FF_READ_ONLY, null);
                        formFields.SetFieldProperty(fieldReasonReason, "setfflags", PdfFormField.FF_READ_ONLY, null);
                    }
                    //Bis zu dieser Stelle sind alle Editierungen des PDFS nur in einem MemoryStream vorhanden. Durch "ToArray()" wird dieser MemoryStream in ein ByteArray umgewandelt.
                    reasonPdf = memoryStream.ToArray();
                }
            }
        }

        public void UpdateListItemReason()
        {
            SPList reasonLibrary = web.Lists[libraryNameReasons];
            reasonFile = web.GetFile(string.Format("{0}/{1}", reasonLibrary.RootFolder.Url, pdfFileNameReason));
            listItemReason = reasonFile.Item;
            //Eintragen der Auftragsnummer in entsprechende Spalte
            listItemReason[collumnReasonOrderNumber] = orderNumber;
            //Ohne SystemUpdate() wird der Spaltenwert nicht aktualisiert. SystemUpdate() ist eine "silent"-Änderung (dh. Änderungsdatum und Editor werden nicht erfasst). Wenn nur "Update()" verwendet wird, werden diese erfasst.
            listItemReason.SystemUpdate();
        }

        //Methode: SharePoint-Designer-Workflow wird gestartet.
        public void StartSPWorkflow(String workflowName)
        {
            SPWorkflowAssociationCollection associationCollection = listItemOrder.ParentList.WorkflowAssociations;
            //Mit foreach wird durch alle vorhandenen Workflows in der Websitesammlung gegangen
            foreach (SPWorkflowAssociation association in associationCollection)
            {
                //Wenn einer der Workflownamen dem Angegebenen (dh. <string workflow>) entspricht wird der enthaltene Code ausgeführt
                if (association.Name == workflowName)
                {
                    //Überprüfen ob eine weitere Instanz des gleichen Workflows schon ausgeführt wird.
                    //Das ist nötig, wenn eine Drittlandsbestellung gemacht wird. Dabei wird der Workflow solange angehalten/pausiert bis die Zollbeauftragten den Auftrag genehmigt haben.
                    //Wenn vor der Genehmigung etwas am Auftrag geändert wird, muss der SP-Designer-WF erneut gestartet werden. Da die vorige Instanz des Workfows aber noch läuft und nur pausiert ist,
                    //schlägt das Starten des Workflows fehl.
                    foreach (SPWorkflow spworkflow in site.WorkflowManager.GetItemActiveWorkflows(listItemOrder))
                    {
                        //Überprüft wird, ob die AssocationId übereinstimmt --> wenn ja, wird diese Instanz mit "CancelWorkflow" abgebrochen
                        if (spworkflow.AssociationId == association.Id) SPWorkflowManager.CancelWorkflow(spworkflow);
                    }
                    association.AutoStartChange = true;
                    association.AutoStartCreate = false;
                    association.AssociationData = string.Empty;
                    //Befehl zum Starten des Workflows
                    site.WorkflowManager.StartWorkflow(listItemOrder, association, association.AssociationData);
                }
            }
        }

        //Methode: Die Methode wird aufgerufen, wenn im Formular enthaltene Anlagen zur Weitergabe markiert wurden. Im Hauptprogramm wird überprüft, 
        //ob das Formularfeld "WeitergabeÄnderung" auf WAHR gesetzt ist. Wenn das der Fall ist, wird KEIN neuer Auftragszettel erstellt, sondern nur diese Methode ausgeführt.
        //Erklärung: Wenn die VW bestimmte Anlagen zur Weitergabe auswählt, ist die Unterschrift des PL schon vorhanden. Durch die Auswahl der Anlagen, wird das Formular geändert,
        //was dazu führt, dass dieser Event Receiver ausgelöst wird (durch das Event "Item updated"). Da in diesem Fall genau derselbe Code ausgeführt werden würde,
        //wie wenn ein neuer Auftrag erstellt worden wäre, würde der Auftragszettel neu erstellt und überschrieben werden. Somit würden alle Signaturen gelöscht werden.
        //Deswegen muss bei Änderung des Formulars geprüft werden, ob es sich um eine inhaltliche Änderung des Auftrags handelt, oder ob "nur" Anlagen zur Weitergabe markiert wurden.
        public void AttachmentsToSupplier()
        {
            xIteratorAttachments = root.Select(xAttachments, nsmgr);
            //"attachments" ist die wiederholte Tabelle mit den Anlagen. Mit "MoveNext()" wird jede Zeile einzelnt bearbeitet.
            while (xIteratorAttachments.MoveNext())
            {
                try
                {
                    //Prüfen, ob das Feld in dem die Anlage "liegt" auch wirklich etwas beinhaltet
                    if (xIteratorAttachments.Current.SelectSingleNode(xAttachment, nsmgr).Value != null)
                    {
                        InfoPathAttachmentDecoder decoder = new InfoPathAttachmentDecoder(xIteratorAttachments.Current.SelectSingleNode(xAttachment, nsmgr).Value);
                        //Laden des Dateinamens der Anlage OHNE die Dateiendung.
                        //Warum? Wenn Word-Dokumente als Anlagen hochgeladen werden, werden diese zu PDF konviertiert. Dadurch ändert sich die Dateiendung. Darum muss "dateiendungs-übergreifend" gearbeitet werden.
                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(decoder.Filename);
                        //SPQuery ist eine Abfrage; anhand dieser können Elemente in Listen und Bibliotheken in Sharepoint gefunden werden
                        SPQuery query = new SPQuery();
                        //Abfrage wird auf rekursiv gestllt, damit sie auch Unterordner erfasst
                        query.ViewAttributes = "Scope=\"Recursive\"";
                        //Eigentlicher Abfrage-String: Es wird abgefragt, ob ein Element vorhanden ist, wessen Dateiname ('FileLeafRef') dem aktuellen Dateiname (fileNamewithoutextension) enthält.
                        //'FileLeafRef' ist ein SharePoint-spezifischer-Name, der den Dateinamen MIT Dateiendung enthält; da wir nur nach Dateiname OHNE Dateiendung abfragen können, wird "Contains" verwendet.
                        query.Query = "<Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Text'>" + fileNameWithoutExtension + "</Value></Contains></Where>";
                        //Angabe des Unterordners in Bibliothek Temp (/temp/<Auftragsnummer>), um die Abfrage nur auf die Anlagen des betroffenen Auftrags auszuführen
                        query.Folder = tempFolder;
                        //Mit diesem Befehl wird die zuvor erstellte Abfrage ausgeführt und das Ergebnis, nämlich alle zutreffenden Elemente, in "collListItems" abgespeichert.
                        //SPListItemCollection ist eine Sammlung von SharePoint-Listenelementen, die später mit foreach weiterverarbeitet werden kann.
                        SPListItemCollection collListItems = libraryTemp.GetItems(query);
                        foreach (SPListItem oListItem in collListItems)
                        {
                            //Überprüfen ob aktuell ausgewählte Anlage zur Weitergabe markiert ist
                            if (xIteratorAttachments.Current.SelectSingleNode(xAttachmentsToSupplier, nsmgr).ValueAsBoolean == true)
                            {
                                //Setzen des Werts der Spalte "Weitergabe" der entsprechenden Datei in Bibliothek "Temp" auf "true"
                                oListItem[collumnTempToSupplier] = true;
                            }
                            else
                            {
                                oListItem[collumnTempToSupplier] = false;
                            }
                            //Updaten des Elements damit Änderungen übernommen werden. SystemUpdate() ist eine "silent"-Änderung (dh. Änderungsdatum und Editor werden nicht erfasst). Wenn nur "Update()" verwendet wird, werden diese erfasst.
                            oListItem.SystemUpdate();
                        }
                    }
                }
                catch { }
            }
        }

        //Mit dieser Methode wird aus den Spalten "Name Besteller" und "Projektleiter" der "SPUser" erhalten; mit diesem kann auf Mail-Adresse und Anzeigename zugegriffen werden
        //Sie wird verwendet, um den Anzeigename in der Mail an die Zentrale verwenden zu können
        public String GetUserDisplayName(SPListItem listItem, String collumnName)
        {
            //Laden des userFields aus dem übergebenem Element (listitem) und der Spalte, die einen Nutzer enthält (spalte)
            SPFieldUser userField = (SPFieldUser)listItem.Fields.GetField(collumnName);
            SPFieldUserValue userFieldValue = (SPFieldUserValue)userField.GetFieldValue(listItem[collumnName].ToString());
            SPUser user = userFieldValue.User;
            //Zuweisen des Anzeigenamens und Rückgabe desselben
            string displayName = user.Name;
            return displayName;
        }
        public byte[] GetFileFromWeb(String folder, String fileName, SPWeb web)
        {
            byte[] file = new byte[0];
            SPList tempList = web.Lists[folder];
            file = web.GetFile(string.Format("{0}/{1}", tempList.RootFolder.Url, fileName)).OpenBinary();
            //file = web.Folders[folder].Files[fileName].OpenBinary();
            return file;
        }
    }
}
