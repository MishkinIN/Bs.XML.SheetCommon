using Bs.XML.SpreadSheet.Resources;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Bs.XML.SpreadSheet {
    internal class SharedStrings {
        public const int MaxCharsInCell = 32767;
        // private const uint defaultCellStyle = 0U;
        // private const uint referenceCellStyle = 1U;
        readonly SharedStringTable sharedStringTable;
        readonly List<string> sharedStrings;
        public SharedStrings(SharedStringTablePart shareStringPart) {
            sharedStrings = new List<string>();
            sharedStringTable = shareStringPart.SharedStringTable;
            if (sharedStringTable is null) {
                sharedStringTable = shareStringPart.SharedStringTable = new SharedStringTable();
                sharedStringTable.Count = sharedStringTable.UniqueCount = 0U;
            }
            foreach (SharedStringItem sharedStringItem in sharedStringTable.Elements<SharedStringItem>()) {
                // Get the text value of the SharedStringItem
                string sharedString = sharedStringItem.InnerText;
                // Add the shared string to the list
                sharedStrings.Add(sharedString);
            }
            sharedStringTable.Count = sharedStringTable.UniqueCount = (uint)sharedStrings.Count;
        }
        public string this[int i] {
            get { return sharedStrings[i]; }
        }
        internal CellValue GetOrCreate(string s) {
            string innerText = s;
            Throw.IfIsTrue(innerText.Length > MaxCharsInCell, 
                () => new Exception(FormattableStringFactory.Create(ExceptionMessage.TooLongCellName, MaxCharsInCell).ToString()));
            int s_index = sharedStrings.IndexOf(innerText);
            if (s_index <= 0) {
                sharedStrings.Add(innerText);
                sharedStringTable.AppendChild<SharedStringItem>(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(innerText)));
                s_index = sharedStrings.Count;
                sharedStringTable.Count = sharedStringTable.UniqueCount = (uint)s_index;
                //shareStrinTable.Save();
            }
            CellValue cv = new CellValue(s_index.ToString());
            return cv;
        }
        internal void Save() {
            sharedStringTable.Save();
        }
    }

}
