using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using DigitalOrdering.Classes.Requirement;

namespace DigitalOrdering.EventReceiver.RequirementFormSubmitted
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class RequirementFormSubmitted : SPItemEventReceiver
    {
        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
            Requirement newRequirement = new Requirement();
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (newRequirement.site = new SPSite(properties.WebUrl))
                {
                    using (newRequirement.web = newRequirement.site.OpenWeb())
                    {
                        newRequirement.listItemForm = newRequirement.web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                        newRequirement.LoadInfoPathForm();
                        newRequirement.CreateTempFolder();
                        newRequirement.GetPdfTemplates();
                        newRequirement.CreateRequirementPdf();
                        newRequirement.UploadFile(newRequirement.libraryNameRequirements, newRequirement.pdfFileNameRequirement, newRequirement.requirementPdf, true);
                        newRequirement.UpdateListItemRequirement();
                        newRequirement.UploadAttachments();
                        if (!newRequirement.reasonEmpty)
                        {
                            newRequirement.CreateReasonPdf();
                            newRequirement.UploadFile(newRequirement.libraryNameReasons, newRequirement.pdfFileNameReason, newRequirement.reasonPdf, true);
                            newRequirement.UpdateListItemReason();
                        }
                        newRequirement.StartSPWorkflow(newRequirement.workflow1Name);
                    }
                }
            });
        }

        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);
            Requirement newRequirement = new Requirement();
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (newRequirement.site = new SPSite(properties.WebUrl))
                {
                    using (newRequirement.web = newRequirement.site.OpenWeb())
                    {
                        newRequirement.listItemForm = newRequirement.web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                        newRequirement.LoadInfoPathForm();
                        newRequirement.CreateTempFolder();
                        newRequirement.GetPdfTemplates();
                        newRequirement.CreateRequirementPdf();
                        newRequirement.UploadFile(newRequirement.libraryNameRequirements, newRequirement.pdfFileNameRequirement, newRequirement.requirementPdf, true);
                        newRequirement.UpdateListItemRequirement();
                        newRequirement.UploadAttachments();
                        if (!newRequirement.reasonEmpty)
                        {
                            newRequirement.CreateReasonPdf();
                            newRequirement.UploadFile(newRequirement.libraryNameReasons, newRequirement.pdfFileNameReason, newRequirement.reasonPdf, true);
                            newRequirement.UpdateListItemReason();
                        }
                        newRequirement.StartSPWorkflow(newRequirement.workflow1Name);
                    }
                }
            });
        }


    }
}