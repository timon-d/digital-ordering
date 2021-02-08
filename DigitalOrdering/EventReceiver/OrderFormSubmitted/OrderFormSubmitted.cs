using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using DigitalOrdering.Classes.Order;

namespace DigitalOrdering.EventReceiver.OrderFormSubmitted
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class OrderFormSubmitted : SPItemEventReceiver
    {
        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);
            Order newOrder = new Order();
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (newOrder.site = new SPSite(properties.WebUrl))
                {
                    using (newOrder.web = newOrder.site.OpenWeb())
                    {
                        newOrder.listItemForm = newOrder.web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                        newOrder.LoadInfoPathForm();
                        if (!newOrder.toSupplierChange)
                        {
                            newOrder.CreateTempFolder();
                            newOrder.GetPdfTemplates();
                            newOrder.CreateOrderPdf();
                            newOrder.UploadFile(newOrder.libraryNameOrders, newOrder.pdfFileNameOrder, newOrder.orderPdf, true);
                            newOrder.UpdateListItemOrder();
                            newOrder.UploadAttachments();
                            if (!newOrder.reasonEmpty)
                            {
                                newOrder.CreateReasonPdf();
                                newOrder.UploadFile(newOrder.libraryNameReasons, newOrder.pdfFileNameReason, newOrder.reasonPdf, true);
                                newOrder.UpdateListItemReason();
                            }
                            newOrder.StartSPWorkflow(newOrder.workflow1Name);
                        }
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
            Order newOrder = new Order();
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (newOrder.site = new SPSite(properties.WebUrl))
                {
                    using (newOrder.web = newOrder.site.OpenWeb())
                    {
                        newOrder.listItemForm = newOrder.web.Lists[properties.ListId].GetItemById(properties.ListItemId);
                        newOrder.LoadInfoPathForm();
                        if (!newOrder.toSupplierChange)
                        {
                            newOrder.CreateTempFolder();
                            newOrder.GetPdfTemplates();
                            newOrder.CreateOrderPdf();
                            newOrder.UploadFile(newOrder.libraryNameOrders, newOrder.pdfFileNameOrder, newOrder.orderPdf, true);
                            newOrder.UpdateListItemOrder();
                            newOrder.UploadAttachments();
                            if (!newOrder.reasonEmpty)
                            {
                                newOrder.CreateReasonPdf();
                                newOrder.UploadFile(newOrder.libraryNameReasons, newOrder.pdfFileNameReason, newOrder.reasonPdf, true);
                                newOrder.UpdateListItemReason();
                            }
                            newOrder.StartSPWorkflow(newOrder.workflow1Name);
                        }
                    }
                }
            });
        }


    }
}