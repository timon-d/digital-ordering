﻿using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace DigitalOrdering.EventReceiver.RequirementPdfDeleting
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class RequirementPdfDeleting : SPItemEventReceiver
    {
        string clmRequirementId = "Title";
        string libNameForms = "BedarfsmeldungFormulare";
        string fileExtensionOrderForm = ".xml";
        string libNameReason = "Begruendungen";
        string fileExtensionReason = "_Begruendung.pdf";
        string libNameTemp = "Temp";
        /// <summary>
        /// An item is being deleted.
        /// </summary>
        public override void ItemDeleting(SPItemEventProperties properties)
        {
            base.ItemDeleting(properties);
            if (properties.ListItem.Name.Contains(".pdf"))
            {
                using (SPSite site = new SPSite(properties.WebUrl))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPListItem _currentItem = web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                        string requirementId = _currentItem[clmRequirementId].ToString();

                        //Wenn der Auftrag ausgecheckt ist, soll das Auschecken verworfen werden
                        SPFile _currentFile = _currentItem.File;
                        if (_currentFile.CheckOutType != SPFile.SPCheckOutType.None)
                        {
                            undoCheckOut(properties);
                        }

                        //Aufragsformular löschen
                        deleteFile(libNameForms, requirementId + fileExtensionOrderForm, web);

                        //Begründung löschen
                        deleteFile(libNameReason, requirementId + fileExtensionReason, web);

                        //Temp-Ordner löschen
                        SPQuery query = new SPQuery();
                        query.Query = "<Where><And><BeginsWith><FieldRef Name='ContentTypeId'/><Value Type='ContentTypeId'>0x0120</Value></BeginsWith><Eq>><FieldRef Name='Title'/><Value Type='Text'>" + requirementId + "</Value></Eq></And></Where>";
                        query.ViewAttributes = "Scope='RecursiveAll'";
                        SPList list = web.Lists[libNameTemp];
                        SPListItemCollection items = list.GetItems(query);
                        foreach (SPListItem i in items)
                        {
                            i.Folder.Recycle();
                        }
                    }
                }
            }
        }
        private void undoCheckOut(SPItemEventProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite(properties.WebUrl))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPFile _currentFile = web.Lists[properties.ListId].GetItemById(properties.ListItemId).File;
                        if (_currentFile.CheckOutType != SPFile.SPCheckOutType.None)
                        {
                            _currentFile.UndoCheckOut();
                        }
                    }
                }
            });
        }
        private void deleteFile(string libName, string fileName, SPWeb web)
        {
                String url = web.Url + "/" + libName + "/" + fileName;
                SPFile file = web.GetFile(url);
                if (file.Exists)
                {
                    file.Item.Recycle();
                }
        }
    }
}