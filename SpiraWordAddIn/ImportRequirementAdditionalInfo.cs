using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace SpiraWordAddIn
{
    /// <summary>
    /// Used to help import formatted text from MS-Word into Spira
    /// </summary>
    class ImportRequirementAdditionalInfo
    {
        public ImportRequirementAdditionalInfo(int exportCount,int errorCount,string requirementName, XmlElement xmlTable)
        {
            this.exportCount = exportCount;
            this.errorCount = errorCount;
            this.requirementName = requirementName;
            this.xmlTable = xmlTable;
        }

        public ImportRequirementAdditionalInfo(int exportCount, int errorCount, string requirementName)
        {
            this.exportCount = exportCount;
            this.errorCount = errorCount;
            this.requirementName = requirementName;
            this.xmlTable = null;
        }


        int exportCount;
        int indentOffset;
        int errorCount;
        string requirementName;
        XmlElement xmlTable;

        public XmlElement XmlTable
        {
            get { return xmlTable; }
            set { xmlTable = value; }
        }

        public int ExportCount
        {
            get { return exportCount; }
            set { exportCount = value; }
        }

        public int ErrorCount
        {
            get { return errorCount; }
            set { errorCount = value; }
        }
        
        public string RequirementName
        {
            get { return requirementName; }
            set { requirementName = value; }
        }
    }
}
