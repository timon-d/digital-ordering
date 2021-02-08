using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    class SharePointUpload
    {
        private byte[] byteArray;
        private String fileName;
        private SPFile newSPFile;
        private String siteUrl;
        private String destinationUrl;
        private SPFolder targetFolder;

        private static String checkInComment = "Check-In durch das System";
        private static Boolean overwriteExistingFile = true;
        
        public SharePointUpload(byte[] byteArray, String fileName, String folderUrl)
        {
            this.byteArray = byteArray;
            this.fileName = fileName;
            this.siteUrl = folderUrl;
            this.targetFolder = GetTargetFolder(folderUrl);
        }

        private SPFolder GetTargetFolder(String folderUrl)
        {
            SPFolder folder;
            using (SPSite site = new SPSite(this.siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    folder = web.GetFolder(folderUrl);
                }
            }
            return folder;
        }

        public String UploadFile()
        {
            using (SPSite site = new SPSite(this.siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    destinationUrl = SPUtility.ConcatUrls(web.Url, targetFolder.Url) + "/" + this.fileName;
                    newSPFile = web.GetFile(destinationUrl);
                    if (newSPFile.Exists && overwriteExistingFile)
                    {
                        if (newSPFile.CheckOutType != SPFile.SPCheckOutType.None)
                        {
                            newSPFile.CheckIn(checkInComment);
                        }
                        newSPFile.Delete();
                    }
                    newSPFile = this.targetFolder.Files.Add(destinationUrl, this.byteArray, overwriteExistingFile);
                }
            }
            return destinationUrl;
        }
    }
}
