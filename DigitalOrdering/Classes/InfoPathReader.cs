using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.XPath;

namespace DigitalOrdering.Classes
{
    class InfoPathReader
    {
        private byte[] xmlFile;
        private XmlDocument xmlDoc;
        private static XPathNavigator root;
        private static XmlNamespaceManager nsmgr;
        private String nameSpacePrefix;
        private String nameSpaceUri;

        class RepeatingTable 
        {
            private Row[] rows;
            private int rowCount;
            private XPathNodeIterator xTableIterator;
            public RepeatingTable(XPathNodeIterator xTableIterator, String[] xPathCollumns)
            {
                this.xTableIterator = xTableIterator;
                this.rowCount = xTableIterator.Count;
                for(int i = 0; i< rowCount; i++)
                {
                    this.rows[i] = new Row(xTableIterator, xPathCollumns);
                }
            }

            public Row[] GetTableRows()
            {
                return this.rows;
            }
        }
        
        class Row 
        {
            private Collumn[] collumns;
            public Row(XPathNodeIterator xTableIterator, String[] xPathCollumns)
            {
                for(int i = 0; i < xPathCollumns.Length; i++)
                {
                    this.collumns[i] = new Collumn(xTableIterator, xPathCollumns[i]);
                }
            }

            public Collumn[] GetCollumns()
            {
                return this.collumns;
            }
        }

        class Collumn
        {
            private XPathNodeIterator xTableIterator;
            private String xPath;
            private XPathNavigator collumnNode;
            private Object collumnValue;
            private Type collumnType;
            public Collumn(XPathNodeIterator xTableIterator, String xPath)
            {
                this.xTableIterator = xTableIterator;
                this.xPath = xPath;
                this.collumnNode = xTableIterator.Current.SelectSingleNode(xPath, nsmgr);

                if (collumnNode.Value != "")
                {
                    TypeCode valueType = Type.GetTypeCode(collumnNode.Value.GetType());
                    if (valueType == TypeCode.Int16 || valueType == TypeCode.Int32 || valueType == TypeCode.Int64)
                    {
                        this.collumnValue = collumnNode.ValueAsInt;
                        this.collumnType = TypeCode.Int32.GetType();
                    }
                    else if (valueType == TypeCode.Double)
                    {
                        this.collumnValue = collumnNode.ValueAsDouble;
                        this.collumnType = TypeCode.Double.GetType();
                    }
                    else if (valueType == TypeCode.DateTime)
                    {
                        this.collumnValue = collumnNode.ValueAsDateTime;
                        this.collumnType = TypeCode.DateTime.GetType();
                    }
                    else if (valueType == TypeCode.Boolean)
                    {
                        this.collumnValue = collumnNode.ValueAsBoolean;
                        this.collumnType = TypeCode.Boolean.GetType();
                    }
                    else
                    {
                        this.collumnValue = collumnNode.Value.ToString();
                        this.collumnType = TypeCode.String.GetType();
                    }
                }
                else
                {
                    this.collumnValue = "";
                    this.collumnType = TypeCode.String.GetType();
                }
            }

            public Object GetCollumnValue()
            {
                return this.collumnValue;
            }
        }

        public InfoPathReader(byte[] xmlFile, String nameSpacePrefix, String nameSpaceUri)
        {
            this.xmlFile = xmlFile;
            this.nameSpacePrefix = nameSpacePrefix;
            this.nameSpaceUri = nameSpaceUri;
            using (Stream xmlMemoryStream = new MemoryStream(this.xmlFile))
            {
                this.xmlDoc = new XmlDocument();
                this.xmlDoc.Load(xmlMemoryStream);
                root = xmlDoc.CreateNavigator();
                nsmgr = new XmlNamespaceManager(new NameTable());
                nsmgr.AddNamespace(this.nameSpacePrefix, this.nameSpaceUri);
            }
        }

        public String GetTextFieldValue(String xPath)
        {
            String value = "";
            XPathNavigator xField = root.SelectSingleNode(xPath, nsmgr);
            if (xField.ToString() != "") value = xField.ToString();
            return value;
        }

        public Boolean GetBooleanFieldValue(String xPath)
        {
            XPathNavigator xField = root.SelectSingleNode(xPath, nsmgr);
            return xField.ValueAsBoolean;
        }

        

    }
}
