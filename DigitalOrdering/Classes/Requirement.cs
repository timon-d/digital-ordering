using InfoPathAttachmentEncoding;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.XPath;


namespace DigitalOrdering.Classes.Requirement
{
    public class Requirement
    {
        //SharePoint
        public SPSite site;
        public SPWeb web;
        public String libraryNameRequirements = "Bedarfsmeldungen";
        public String libraryNameReasons = "Begründungen";
        public String folderNameReasons = "Begruendungen";
        public SPList libraryReasons;
        public String libraryNameTemp = "Temp";
        public SPFolder tempFolder;
        public SPList libraryTemp;
        public SPListItem listItemForm;
        public SPListItem listItemRequirement;
        public SPListItem listItemReason;
        public SPFile requirementFile;
        public SPFile reasonFile;

        public String templateFolder = "Vorlage";
        public String template = "Vorlage_Bedarfsmeldung.pdf";
        public String templateReason = "Vorlage_Bedarfsmeldung_Begruendung.pdf";
        public String checkInComment = "Check-In durch das System";
        public String pdfFileNameRequirement;
        public String pdfFileNameReason;
        public String collumnRequirementNumber = "Title";
        public String collumnFormRequirementCustomer = "NameBesteller";
        public String collumnFormRequirementProjectleader = "Projektleiter";
        public String collumnRequirementDate = "Datum";
        public String collumnRequirementLaser = "Laser";
        public String collumnRequirementDanger = "Gefahrstoff";
        public String collumnRequirementIDDB = "IDDB";
        public String collumnRequirementProject = "Projekt-Nr.";
        public String collumnRequirementAttachments = "Anlagen";
        public String collumnRequirementGroup = "Gruppe";
        public String collumnRequirementApprovalDate = "Datum Zweitunterschrift";
        public String collumnRequirementSubmitCounter = "SubmitCounter";
        public String collumnRequirementInvest = "Invest";

        public String collumnTempTitle = "Title";
        public String collumnFormRequirementId = "BM-Nr.";
        public String collumnReasonOrderNumber = "Auftragsnummer";
        public String workflow1Name = "Bedarfsmeldung - Formular abgeschickt";
        public String workflow2Name = "Bedarfsmeldung: Genehmigungsvorgang";

        //Infopath-Formular 
        public XPathNavigator root;
        public XmlNamespaceManager nsmgr;
        public String nameSpacePrefix = "my";
        public String nameSpaceUri = "http://schemas.microsoft.com/office/infopath/2003/myXSD/2019-02-06T08:35:46";
        public String xRequirementID = "/my:myFields/my:requirementID";
        public String xDeliveryDate = "/my:myFields/my:deliveryDate";
        public String xIsInvest = "/my:myFields/my:isInvest";
        public String xIsDraft = "/my:myFields/my:isDraft";
        public String xCountSubmitted = "/my:myFields/my:countSubmitted";
        public String xDanger = "/my:myFields/my:isDangerousSubstance";
        public String xLaser = "/my:myFields/my:isLaser";
        public String xTelephone = "/my:myFields/my:customerPhoneExtension";
        public String xNote = "/my:myFields/my:note";
        public String xProjectNumber = "/my:myFields/my:projectNumber";
        public String xIsMultipleProjects = "/my:myFields/my:isMultipleProjects";
        public String xMultipleProjectsBookingType = "/my:myFields/my:multipleProjectsBookingType";
        public String xAttachments = "/my:myFields/my:attachments/my:attachment";
        public XPathNodeIterator xIteratorAttachments;
        public String xAttachment = "my:file";
        public String xPositions = "/my:myFields/my:positions/my:position";
        public XPathNodeIterator xIteratorPositions;
        public String xPositionNumber = "my:positionNumber";
        public String xPositionQuantity = "my:productQuantity";
        public String xPositionProjectNumber = "my:areaProjectNumber/my:positionProjectNumber";
        public String xPositionProductDescription = "my:productDescription";
        public String xPositionProductPrice = "my:productPrice";
        public String xPercentageProject = "/my:myFields/my:areaProjectsPercentage/my:ProjectsPercentage/my:Project";
        public XPathNodeIterator xIteratorProjects;
        public String xPercentage = "my:percentage";
        public String xPercentageProjectNumber = "my:projectNumberPercentage";
        public String xAttachmentsCheck = "/my:myFields/my:isAttachment";
        public String xGroup = "/my:myFields/my:groupName";
        public XPathNodeIterator xIteratorCompetitors;
        String xCompetitors = "/my:myFields/my:competitors/my:competitor";
        String xCompetitorName = "my:competitorName";
        String xCompetitorLocation = "my:competitorLocation";
        String xCompetitorCountry = "my:competitorCountry";
        XPathNodeIterator xIteratorCriteria;
        String xCriteria = "/my:myFields/my:awardCriteria/my:criteria";
        String xCriteriaName = "my:criteriaName";
        String xCriteriaImportance = "my:criteriaImportance";
        String xReason = "/my:myFields/my:reason";

        //PDF-Formularfelder
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

        //Sonstiges
        public Font frutigerFont = new Font(BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.CP1252, BaseFont.EMBEDDED), 10);
        public int currencyFormat = 1031;
        public String requirementId;
        public Boolean attachmentsCheck;
        public Boolean isInvest;
        public Boolean reasonEmpty = true;
        public String Invest = "Nein";
        public String danger = "Nein";
        public String laser = "Nein";
        public byte[] requirementTemplate = new byte[0];
        public byte[] requirementPdf = new byte[0];
        public byte[] reasonPdf = new byte[0];
        public String tempUploadUrl;
        string pathToGroupMappingsFile = @"C:\Bestellung\Config\groupMapping.txt";

        public void LoadInfoPathForm()
        {
            byte[] xmlFile = listItemForm.File.OpenBinary();
            Stream xmlMemoryStream = new MemoryStream(xmlFile);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlMemoryStream);
            root = xmlDoc.CreateNavigator();
            nsmgr = new XmlNamespaceManager(new NameTable());
            nsmgr.AddNamespace(nameSpacePrefix, nameSpaceUri);
            requirementId = root.SelectSingleNode(xRequirementID, nsmgr).ToString();
            pdfFileNameRequirement = requirementId + ".pdf";
            pdfFileNameReason = requirementId + "_Begruendung.pdf";
            if (root.SelectSingleNode(xDanger, nsmgr).ToString() != "false") danger = "Ja";
            if (root.SelectSingleNode(xLaser, nsmgr).ToString() != "false") laser = "Ja";
            isInvest = root.SelectSingleNode(xIsInvest, nsmgr).ValueAsBoolean;
            if (isInvest) Invest = "Ja";
            if (root.SelectSingleNode(xReason, nsmgr).ToString() != "") reasonEmpty = false;
            attachmentsCheck = root.SelectSingleNode(xAttachmentsCheck, nsmgr).ValueAsBoolean;
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
            tempFolder = web.GetFolder(String.Format("{0}/Temp/" + requirementId, web.Url));
            if (tempFolder.Exists)
            {
                tempFolder.Delete();
            }
            var i = libraryTemp.Items.Add("", SPFileSystemObjectType.Folder, requirementId);
            //Ohne die Update()-Funktion wird kein Ordner angelegt!
            i.Update();
            tempFolder.Item[collumnTempTitle] = requirementId;
            tempFolder.Item.Update();
            tempUploadUrl = libraryTemp.RootFolder.SubFolders[requirementId].Url + "/";
        }

        //Methode: Es wird eine der PDF-Vorlagen aus der Bibliothek "Vorlage" je nach angegebenen Werten im Auftragsformular geladen und als Byte-Array zurück gegeben.
        //Mit web.folders[<Bibliothek>].Files[<Dateiname>].OpenBinary() wird eine Datei binär bzw. als Byte-Array geöffnet
        public void GetPdfTemplates()
        {
            requirementTemplate = web.Folders[templateFolder].Files[template].OpenBinary();
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
        public void CreateRequirementPdf()
        {
            Dictionary<string, string> groupMapping = new Dictionary<string, string>();
            using (var sr = new StreamReader(pathToGroupMappingsFile))
            {
                string line = null;
                while ((line = sr.ReadLine()) != null)
                {
                    string oe = line.Substring(0, line.IndexOf(','));
                    string group = line.Substring(line.LastIndexOf(',') + 1);
                    groupMapping.Add(oe, group);
                }
            }

            int xDistance = 34;
            int yDistance = 745;
            float rowMinHeight = 15f;
            float currentPos;
            //initialiseren eines MemoryStreams, welcher dem PDF-Stamper (siehe nächster Befehl) übergeben wird
            using (Document document = new Document(PageSize.A4))
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    //Folgende drei Initalisierungen werden in Kombination mit using(){} gemacht. Somit gelten die Initialisierungen nur innerhalb des using-Abschnitts.
                    //Das scheint C#-Best-Practice zu sein, da ansonsten u.a. das bearbeitete PDF für immer in einem "Die Datei wird schon verwendet"-Zustand bleibt und nicht verwendet werden kann.
                    //initialisieren des PDF-Readers von itextsharp (zum Lesen von PDFs); diesem wird die Vorlage des Auftragszettels als ByteArray übergeben
                    using (PdfReader pdfReader = new PdfReader(requirementTemplate))
                    {
                        //initialisieren des PDF-Stampers von itextsharp (zum Bearbeiten von PDFs)
                        using (PdfStamper pdfStamper = new PdfStamper(pdfReader, memoryStream, '\0', true))
                        {
                            PdfContentByte canvas = pdfStamper.GetOverContent(1);

                            ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase("Bedarfsmeldung #" + requirementId), xDistance, yDistance, 0);

                            PdfPTable pdfInfoTable = new PdfPTable(2);
                            pdfInfoTable.SetTotalWidth(new float[] { 100, 400 });
                            AddPDFCell("Invest: ", 0, 2, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(Invest, 0, 2, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Projektnummer: ", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            string projectNumber = "";
                            if (root.SelectSingleNode(xIsMultipleProjects, nsmgr).ToString() != "false")
                            {
                                if (root.SelectSingleNode(xMultipleProjectsBookingType, nsmgr).Value == "percentage")
                                {
                                    xIteratorProjects = root.Select(xPercentageProject, nsmgr);
                                    while (xIteratorProjects.MoveNext()) 
                                    {
                                        projectNumber += xIteratorProjects.Current.SelectSingleNode(xPercentageProjectNumber, nsmgr).Value + " (" + xIteratorProjects.Current.SelectSingleNode(xPercentage, nsmgr).Value + "%)\n";
                                    }
                                }
                            }
                            else
                            {
                                projectNumber = root.SelectSingleNode(xProjectNumber, nsmgr).ToString();
                            }
                            AddPDFCell(projectNumber, 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Name Besteller:", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(listItemForm[collumnFormRequirementCustomer].ToString(), 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Durchwahl Besteller:", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(root.SelectSingleNode(xTelephone, nsmgr).ToString(), 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Projektleiter:", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(listItemForm[collumnFormRequirementProjectleader].ToString(), 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Gruppe/Abteilung:", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(groupMapping[root.SelectSingleNode(xGroup, nsmgr).ToString()], 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Laser:", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(laser, 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell("Gefahrstoff:", 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            AddPDFCell(danger, 0, 1, 0, 0, rowMinHeight, pdfInfoTable);
                            currentPos = pdfInfoTable.WriteSelectedRows(0, 2, 0, 8, xDistance, yDistance - 20, pdfStamper.GetOverContent(1));

                            PdfPTable pdfTablePositions = new PdfPTable(4);
                            pdfTablePositions.SetTotalWidth(new float[] { 28, 40, 330, 128 });
                            pdfTablePositions.DefaultCell.Border = Rectangle.NO_BORDER;
                            xIteratorPositions = root.Select("/my:myFields/my:positions/my:position", nsmgr);
                            AddPDFCell("Pos", 0, 2, 0, 0, 20f, pdfTablePositions);
                            AddPDFCell("Menge", PdfPCell.ALIGN_CENTER, 2, 4, 0, 20f, pdfTablePositions);
                            AddPDFCell("Warenbezeichnung", 0, 2, 4, 0, 20f, pdfTablePositions);
                            AddPDFCell("Geschätzter Auftragswert", PdfPCell.ALIGN_RIGHT, 2, 4, 0, 20f, pdfTablePositions);
                            pdfTablePositions.HeaderRows = 1;
                            while (xIteratorPositions.MoveNext())
                            {
                                AddPDFCell(xIteratorPositions.Current.SelectSingleNode(xPositionNumber, nsmgr).Value, 0, 0, 0, 0, rowMinHeight, pdfTablePositions);
                                AddPDFCell(xIteratorPositions.Current.SelectSingleNode(xPositionQuantity, nsmgr).Value, PdfPCell.ALIGN_CENTER, 4, 4, 0, rowMinHeight, pdfTablePositions);
                                string sonstiges = "";
                                if (xIteratorPositions.Current.SelectSingleNode(xPositionProjectNumber, nsmgr).Value != String.Empty)
                                {
                                    sonstiges = sonstiges + "\nProjekt-Nr.: " + xIteratorPositions.Current.SelectSingleNode(xPositionProjectNumber, nsmgr).Value;
                                }
                                AddPDFCell(xIteratorPositions.Current.SelectSingleNode(xPositionProductDescription, nsmgr).Value + sonstiges, 0, 4, 4, 0, rowMinHeight, pdfTablePositions);
                                if (xIteratorPositions.Current.SelectSingleNode(xPositionProductPrice, nsmgr).Value.ToString() != "")
                                {
                                    AddPDFCell(CurrencyToString(xIteratorPositions.Current.SelectSingleNode(xPositionProductPrice, nsmgr).ValueAsDouble), PdfPCell.ALIGN_RIGHT, 4, 4, 0, rowMinHeight, pdfTablePositions);
                                }
                            }
                            currentPos = pdfTablePositions.WriteSelectedRows(0, 4, 0, xIteratorPositions.Count + 1, xDistance, currentPos - 30, pdfStamper.GetOverContent(1));

                            PdfPTable pdfTableCompetitors = new PdfPTable(3);
                            pdfTableCompetitors.SetTotalWidth(new float[] { 306, 110, 110 });
                            pdfTableCompetitors.DefaultCell.Border = Rectangle.NO_BORDER;
                            AddPDFCell("Wettbewerber", 0, 2, 0, 0, 20f, pdfTableCompetitors);
                            AddPDFCell("Ort", 0, 2, 5, 0, 20f, pdfTableCompetitors);
                            AddPDFCell("Land", 0, 2, 5, 0, 20f, pdfTableCompetitors);
                            pdfTableCompetitors.HeaderRows = 1;
                            xIteratorCompetitors = root.Select(xCompetitors, nsmgr);
                            while (xIteratorCompetitors.MoveNext())
                            {
                                AddPDFCell(xIteratorCompetitors.Current.SelectSingleNode(xCompetitorName, nsmgr).Value, 0, 0, 0, 0, rowMinHeight, pdfTableCompetitors);
                                AddPDFCell(xIteratorCompetitors.Current.SelectSingleNode(xCompetitorLocation, nsmgr).Value, 0, 4, 5, 0, rowMinHeight, pdfTableCompetitors);
                                AddPDFCell(xIteratorCompetitors.Current.SelectSingleNode(xCompetitorCountry, nsmgr).Value, 0, 4, 5, 0, rowMinHeight, pdfTableCompetitors);
                            }
                            currentPos = pdfTableCompetitors.WriteSelectedRows(0, 3, 0, xIteratorCompetitors.Count + 1, xDistance, currentPos - 30, pdfStamper.GetOverContent(1));

                            PdfPTable pdfTableCriteria = new PdfPTable(2);
                            pdfTableCriteria.SetTotalWidth(new float[] { 150, 45 });
                            pdfTableCriteria.DefaultCell.Border = Rectangle.NO_BORDER;
                            AddPDFCell("Kriterium", 0, 2, 0, 0, 20f, pdfTableCriteria);
                            AddPDFCell("%", 0, 2, 4, 0, 20f, pdfTableCriteria);
                            pdfTableCriteria.HeaderRows = 1;
                            xIteratorCriteria = root.Select(xCriteria, nsmgr);
                            while (xIteratorCriteria.MoveNext())
                            {
                                AddPDFCell(xIteratorCriteria.Current.SelectSingleNode(xCriteriaName, nsmgr).Value, 0, 0, 0, 0, rowMinHeight, pdfTableCriteria);
                                AddPDFCell(xIteratorCriteria.Current.SelectSingleNode(xCriteriaImportance, nsmgr).Value, 0, 4, 4, 0, rowMinHeight, pdfTableCriteria);
                            }
                            currentPos = pdfTableCriteria.WriteSelectedRows(0, 2, 0, xIteratorCriteria.Count + 1, xDistance, currentPos - 30, pdfStamper.GetOverContent(1));

                            PdfPTable pdfTableInfo2 = new PdfPTable(2);
                            pdfTableInfo2.SetTotalWidth(new float[] { 140, 386 });
                            pdfTableInfo2.DefaultCell.Border = Rectangle.NO_BORDER;
                            AddPDFCell("Bedarfstermin: ", 0, 0, 0, 0, rowMinHeight, pdfTableInfo2);
                            if (root.SelectSingleNode(xDeliveryDate, nsmgr).Value != "")
                            {
                                var date = Convert.ToDateTime(root.SelectSingleNode(xDeliveryDate, nsmgr).Value);
                                AddPDFCell(date.ToString("dd.MM.yyyy"), 0, 0, 0, 0, rowMinHeight, pdfTableInfo2);
                            }
                            else {
                                AddPDFCell("", 0, 0, 0, 0, rowMinHeight, pdfTableInfo2);
                            }
                            AddPDFCell("Anmerkung für den Einkauf: ", 0, 0, 0, 0, rowMinHeight, pdfTableInfo2);
                            AddPDFCell(root.SelectSingleNode(xNote, nsmgr).Value, 0, 0, 0, 0, rowMinHeight, pdfTableInfo2);
                            currentPos = pdfTableInfo2.WriteSelectedRows(0, 2, 0, 2, xDistance, currentPos - 30, pdfStamper.GetOverContent(1));
                            ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase("Unterschrift Projektleiter"), xDistance, 80f, 0);
                        }
                    }
                    requirementPdf = memoryStream.ToArray();
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

        public void UpdateListItemRequirement()
        {
            requirementFile = web.Folders[libraryNameRequirements].Files[pdfFileNameRequirement];
            listItemRequirement = requirementFile.Item;
            listItemRequirement[collumnRequirementNumber] = requirementId;   //String
            listItemRequirement[collumnRequirementIDDB] = listItemForm.ID.ToString();   //String
            listItemRequirement[collumnRequirementDanger] = danger; //String
            listItemRequirement[collumnRequirementLaser] = laser;  //String
            listItemRequirement[collumnRequirementProject] = root.SelectSingleNode(xProjectNumber, nsmgr).Value.ToString(); //String
            listItemRequirement[collumnRequirementDate] = listItemForm["Erstellt"]; //Datum
            listItemRequirement[collumnRequirementApprovalDate] = DateTime.Now;
            listItemRequirement[collumnRequirementGroup] = web.SiteGroups[root.SelectSingleNode(xGroup, nsmgr).Value.ToString()];  //Personenfeld    
            int submitCounter = Int32.Parse(listItemRequirement[collumnRequirementSubmitCounter].ToString());
            submitCounter++;
            listItemRequirement[collumnRequirementSubmitCounter] = submitCounter;
            listItemRequirement[collumnRequirementInvest] = isInvest;
            if (attachmentsCheck)
            {
                SetHyperlinkField(collumnRequirementAttachments, "öffnen", "Attachments");
            }
            else
            {
                listItemRequirement[collumnRequirementAttachments] = null;
            }
            listItemRequirement.SystemUpdate();
        }

        private SPFieldUserValue getSPFieldUserValue(String userName)
        {
            SPUser user = web.EnsureUser(userName);
            SPFieldUserValue userValue = new SPFieldUserValue(web, user.ID, user.LoginName);
            return userValue;
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
            listItemRequirement[collumn] = urlFieldValue;
        }

        //Methode: Decodieren der Anlagen und Speichern in Bibliothek "Temp" im entsprechenden Unterordner
        public void UploadAttachments()
        {
            xIteratorAttachments = root.Select(xAttachments, nsmgr);
            //Durch jede Zeile der wiederholten Tabelle der Anlagen iterieren mit MoveNext
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

        public void CreateReasonPdf()
        {
            int xDistance = 72;
            int yDistance = 745;
            float rowMinHeight = 15f;
            float currentPos;
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
                        PdfContentByte canvas = pdfStamper.GetOverContent(1);
                        ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, new Phrase("Begründung #" + requirementId), xDistance, yDistance -30, 0);
                        PdfPTable pdfTable = new PdfPTable(1);
                        pdfTable.SetTotalWidth(new float[] { 450 });
                        pdfTable.DefaultCell.Border = Rectangle.NO_BORDER;
                        AddPDFCell(root.SelectSingleNode(xReason, nsmgr).ToString(), 0, 0, 0, 0, rowMinHeight, pdfTable);
                        currentPos = pdfTable.WriteSelectedRows(0, 2, 0, 2, xDistance, yDistance - 50, pdfStamper.GetOverContent(1));
                    }
                    //Bis zu dieser Stelle sind alle Editierungen des PDFS nur in einem MemoryStream vorhanden. Durch "ToArray()" wird dieser MemoryStream in ein ByteArray umgewandelt.
                    reasonPdf = memoryStream.ToArray();
                }
            }
        }

        public void UpdateListItemReason()
        {
            //reasonFile = web.Folders[folderNameReasons].Files[pdfFileNameReason];
            SPList reasonLibrary = web.Lists[libraryNameReasons];
            reasonFile = web.GetFile(string.Format("{0}/{1}", reasonLibrary.RootFolder.Url, pdfFileNameReason));
            listItemReason = reasonFile.Item;
            //Eintragen der Auftragsnummer in entsprechende Spalte
            listItemReason[collumnReasonOrderNumber] = requirementId;
            //Ohne SystemUpdate() wird der Spaltenwert nicht aktualisiert. SystemUpdate() ist eine "silent"-Änderung (dh. Änderungsdatum und Editor werden nicht erfasst). Wenn nur "Update()" verwendet wird, werden diese erfasst.
            listItemReason.SystemUpdate();
        }

        //Methode: SharePoint-Designer-Workflow wird gestartet.
        public void StartSPWorkflow(String workflowName)
        {
            SPWorkflowAssociationCollection associationCollection = listItemRequirement.ParentList.WorkflowAssociations;
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
                    foreach (SPWorkflow spworkflow in site.WorkflowManager.GetItemActiveWorkflows(listItemRequirement))
                    {
                        //Überprüft wird, ob die AssocationId übereinstimmt --> wenn ja, wird diese Instanz mit "CancelWorkflow" abgebrochen
                        if (spworkflow.AssociationId == association.Id) SPWorkflowManager.CancelWorkflow(spworkflow);
                    }
                    association.AutoStartChange = true;
                    association.AutoStartCreate = false;
                    association.AssociationData = string.Empty;
                    //Befehl zum Starten des Workflows
                    site.WorkflowManager.StartWorkflow(listItemRequirement, association, association.AssociationData);
                }
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
