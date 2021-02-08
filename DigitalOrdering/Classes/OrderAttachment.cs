using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    class OrderAttachment
    {
        private int attachmentNumber;
        private String fileName;
        private String fileExtension;
        private Boolean isForSupplier;
        private String fileUrl;
        private byte[] byteArray;

        public OrderAttachment(int attachmentNumber, String fileName, Boolean isForSupplier)
        {
            this.attachmentNumber = attachmentNumber;
            this.fileName = fileName;
            
            this.isForSupplier = isForSupplier;
        }

        public void UploadAttachment(String folderUrl)
        {
            SharePointUpload spUpload = new SharePointUpload(this.byteArray, this.fileName, folderUrl);
            this.fileUrl = spUpload.UploadFile();
        }
    }
}
