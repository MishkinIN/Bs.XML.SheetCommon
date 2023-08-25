using Bs.XML.SpreadSheet.Resources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Bs.XML.SpreadSheet {
    /// <summary>
    /// Предоставляет базовые операции по использованию Excell файла как ресурса.
    /// Для файла ресурса устанавливается "Действие при сборке: Нет", "Копировать в выходной каталог: Всегда копировать".
    /// </summary>
    public abstract class XlsxResource : IDisposable {
        private bool disposed = false;
        private SharedStringItem[] stringValues = Array.Empty<SharedStringItem>();
        private readonly Lazy<SpreadsheetDocument> lSpreadsheetDocument;
        private readonly bool forWrite;
        protected string sheetName;
        private bool disposedValue;

        /// <summary>
        /// Имя ресурса, включая путь относительно выполняемой сборки. 
        /// </summary>
        /// <example>Resources\MyResource</example>
        public string Resource { get; protected set; }
        protected XlsxResource(string fileName, bool forWrite) {
            sheetName = string.Empty;
            lSpreadsheetDocument = new Lazy<SpreadsheetDocument>(() => LoadSpreadSheet(fileName, forWrite));
            Resource = fileName;
            this.forWrite = forWrite;
        }
        protected XlsxResource(Stream stream) {
            sheetName = string.Empty;
            Resource = string.Empty;
            this.forWrite = false;
            SpreadsheetDocument spreadsheetDocument = LoadSpreadSheet(stream, false);
            lSpreadsheetDocument = new Lazy<SpreadsheetDocument>(() => spreadsheetDocument);
        }
        protected string GetStringValue(Cell cell) {
            if (cell.CellValue == null)
                return String.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var index = int.Parse(cell.CellValue.Text);
                var value = (StringValues[index].InnerText).Trim();
                return value;
            }
            else {
                return cell.CellValue.Text;
            }
        }

        protected int? GetIntValue(Cell cell) {
            if (cell.CellValue == null)
                return null;
#pragma warning disable IDE0018 // Объявление встроенной переменной
            int value;
#pragma warning restore IDE0018 // Объявление встроенной переменной
            if (int.TryParse(GetStringValue(cell), NumberStyles.Number, CultureInfo.InvariantCulture, out value))
                return value;
            else
                return null;
        }
        protected double? GetDoubleValue(Cell cell) {
            if (cell.CellValue == null)
                return null;
            double value;
            string sValue = GetStringValue(cell);
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                if (double.TryParse(sValue, NumberStyles.Number, CultureInfo.CurrentCulture, out value))
                    return value;
            }
            else
            if (double.TryParse(sValue, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
                return value;
            return null;
        }

        protected SheetData GetSheetData() {
            var thisSheetName = String.IsNullOrEmpty(sheetName) ? "Лист1" : sheetName;
            WorkbookPart? wbPart = lSpreadsheetDocument.Value?.WorkbookPart;
            Throw.IfNull(wbPart, () => throw new InvalidOperationException(ExceptionMessage.XlsxResourceNotLoaded));
            Sheet? sh = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => sheetName.Equals(s.Name));
            if (sh is null) {
                sh = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                Throw.IfNull(sh,()=> throw new InvalidOperationException(ExceptionMessage.WorkbookHasNotAnySheet));
                sheetName = sh.Name!.Value ?? string.Empty;
            }
            StringValue relId = sh.Id!;
            WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(relId.Value!);
            //var sharedStringPart = wbPart.SharedStringTablePart;

            SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();
            return sheetData;
        }
        protected void SetValue(Cell outCell, string text) {
            int index = InsertSharedStringItem(text, this.GetSharedStringTablePart());
            outCell.CellValue = new CellValue(index.ToString());
            outCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
            if (index >= stringValues.Length) {
                RefreshStringValues();
            }
        }
        protected SharedStringTablePart GetSharedStringTablePart() {
            SharedStringTablePart shareStringPart;
            if (lSpreadsheetDocument.Value.WorkbookPart!.GetPartsOfType<SharedStringTablePart>().Count() > 0) {
                shareStringPart = lSpreadsheetDocument.Value.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else {
                shareStringPart = lSpreadsheetDocument.Value.WorkbookPart.AddNewPart<SharedStringTablePart>();

            }
            return shareStringPart;
        }
        protected void SaveFile() {
            if (lSpreadsheetDocument.IsValueCreated) {
                lSpreadsheetDocument.Value.Save();
            }
        }
        protected int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart) {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null) {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>()) {
                if (item.InnerText == text) {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();
            stringValues = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            return i;
        }
        private SharedStringItem[] StringValues {
            get {
                if (stringValues == null) {
                    RefreshStringValues();
                }
                return stringValues!;
            }
        }

        private void RefreshStringValues() {
            WorkbookPart wbPart = lSpreadsheetDocument.Value.WorkbookPart!;
            var shareStringPart = wbPart.SharedStringTablePart;
            if (shareStringPart is null)
                stringValues = Array.Empty<SharedStringItem>();
            else {
                stringValues = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            }
        }

        private static SpreadsheetDocument LoadSpreadSheet(string fileFullName, bool forWrite) {
            return SpreadsheetDocument.Open(fileFullName, isEditable: forWrite);
        }
        private static SpreadsheetDocument LoadSpreadSheet(Stream stream, bool forWrite) {
            return SpreadsheetDocument.Open(stream, isEditable: forWrite);
        }

        protected virtual void Dispose(bool disposing) {
            if (!disposedValue) {
                if (disposing) {
                    if (lSpreadsheetDocument.IsValueCreated) {
                        lSpreadsheetDocument.Value.Dispose();
                    }
                }
                // // set large fields to null
                disposedValue = true;
            }
        }

        public void Dispose() {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
