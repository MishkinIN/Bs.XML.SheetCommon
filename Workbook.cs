using Bs.XML.SpreadSheet.Resources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace Bs.XML.SpreadSheet {
    /// <summary>
    /// Предоставляет специализированные операции по работе с файлом Excell
    /// </summary>
    internal class Workbook : IDisposable {
        private readonly FileInfo file;
        private readonly bool disposed = false;
        private SpreadsheetDocument? spreadsheetDocument = null;
        //private SharedStringItem[]? stringValues = null;
        private SharedStrings? sharedStrings = null;
        private bool disposedValue;
        private readonly bool isEditable = false;
        private IDisposable? stream = null;

        internal Workbook(string name) : this(new FileInfo(name)) {
        }
        internal Workbook(FileInfo fileInfo, bool isEditable = false) {
            this.file = fileInfo;
            this.isEditable = isEditable;
        }
        internal static Workbook FromTemplate(string templateFullName, string targetName) {
            FileInfo fiTemplate = new FileInfo(templateFullName);
            Throw.IfIsTrue(!fiTemplate.Exists,
                () => new FileNotFoundException(FormattableStringFactory.Create(ExceptionMessage.TemplateFileNameNotFound, templateFullName).ToString()));
            FileInfo fiTarget = new FileInfo(targetName);
            using (var sourcePackage = SpreadsheetDocument.Open(fiTemplate.FullName, isEditable: false)) {
                // Create a new package and clone the content from the source package
                var wb = new Workbook(fiTarget, isEditable: true);
                FileStream stream = fiTarget.Open(FileMode.Create, FileAccess.ReadWrite);
                wb.spreadsheetDocument = (SpreadsheetDocument)(sourcePackage.Clone(stream, isEditable: true, openSettings: new OpenSettings { AutoSave = true }));
                wb.stream = stream;
                WorkbookPart workbookPart = wb.spreadsheetDocument.WorkbookPart!;
                SharedStringTablePart sharedStringsPart = workbookPart.SharedStringTablePart
                    ?? workbookPart.AddNewPart<SharedStringTablePart>();
                wb.sharedStrings = new SharedStrings(sharedStringsPart);
                return wb;
            }

        }
        internal string FileName => file.FullName;

        internal Row? Row(WorksheetPart worksheetPart, UInt32 rowIndex) {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            Throw.IfNull(sheetData,()=> FormattableStringFactory.Create(ExceptionMessage.SheetDataNotFound));
            return GetRows(sheetData).FirstOrDefault(r => ((UInt32)r.RowIndex!) == rowIndex);
        }
        internal Row GetOrAddRow(WorksheetPart worksheetPart, UInt32 rowIndex) {
            Row? row = Row(worksheetPart, rowIndex);
            if (row is null) {
                SheetData? sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                Throw.IfNull(sheetData, () => FormattableStringFactory.Create(ExceptionMessage.SheetDataNotFound));
                row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            return row;
        }
        internal Cell FindCell(Row row, ColumnName column) {
            string cellReference = column.Name + row.RowIndex;
            // If there is not a cell with the specified column name, insert one.  
            var cell= row.Elements<Cell>().FirstOrDefault(c=>c.CellReference== cellReference);
            Throw.IfNull(cell, () => new KeyNotFoundException());
            return cell;
        }
        internal Cell GetOrAddCell(WorksheetPart worksheetPart, Row row, ColumnName column) {
            {
                //Worksheet worksheet = worksheetPart.Worksheet;
                //SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                string cellReference = column.Name + row.RowIndex;

                // If there is not a cell with the specified column name, insert one.  
                var refCell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);
                if (refCell is not null) {
                    return refCell;
                }
                else {
                    // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                    refCell = row.Elements<Cell>().FirstOrDefault(cell => string.Compare(cell.CellReference!.Value, cellReference, true) > 0);
                    Cell newCell = new Cell() { CellReference = cellReference };
                    if (refCell is null) { 
                        row.AppendChild(newCell);
                    }
                    else {
                        row.InsertBefore(newCell, refCell); 
                    }
                    return newCell;
                }
            }
        }
        internal IEnumerable<Row> GetRows(SheetData sheetData) {
            foreach (Row row in sheetData.Elements<Row>()) {
                yield return row;
            }
        }
        internal IEnumerable<Cell> GetCells(Row row) {
            foreach (Cell cell in row.Elements<Cell>()) {
                yield return cell;
            }
        }
        internal WorksheetPart GetWorksheetPart(string name) {
            WorkbookPart wbPart = SpreadsheetDoc.WorkbookPart!;
            string? relId = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => String.Equals(name, s.Name))?.Id;
            Throw.IfNull(relId,()=>new InvalidOperationException( FormattableStringFactory.Create(ExceptionMessage.SheetNameNotFound,name).ToString()));
            WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(relId);
            return wsPart;
        }
        internal SheetData? GetSheetData(string sheetName) {
            WorksheetPart wsPart;
            WorkbookPart wbPart = SpreadsheetDoc.WorkbookPart!;
            string relId;
            if (String.IsNullOrEmpty(sheetName)) {
                Sheet? sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet is null)
                    return null;
                relId = sheet.Id!;
            }
            else {
                Sheet? sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => sheetName.Equals(s.Name));
                if (sheet is null)
                    return null;
                relId = sheet.Id!;
            }
            wsPart = (WorksheetPart)wbPart.GetPartById(relId);
            SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();
            return sheetData;
        }
        internal string GetStringValue(Cell cell) {
            if (cell.CellValue == null)
                return String.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var index = int.Parse(cell.CellValue.Text);
                var value = sharedStrings![index];
                return value;
            }
            else {
                return cell.CellValue.Text;
            }
        }
        internal int? GetIntValue(Cell cell) {
            if (cell.CellValue == null)
                return null;
            if (int.TryParse(GetStringValue(cell), NumberStyles.Number, CultureInfo.InvariantCulture, out int value))
                return value;
            else
                return null;
        }
        internal double? GetDoubleValue(Cell cell) {
            if (cell.CellValue == null)
                return null;
            if (double.TryParse(GetStringValue(cell), NumberStyles.Number, CultureInfo.InvariantCulture, out double value))
                return value;
            else
                return null;
        }
        internal void SetValue(Cell outCell, string text) {
            if (string.IsNullOrEmpty(text)) {
                SetCellEmptyValue(outCell);
                return;
            }
            CellValue textRef = sharedStrings!.GetOrCreate(text);
            outCell.CellValue = textRef;
            outCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        }

        internal void SetValue(Cell cell, int? data) {
            if (data.HasValue) {
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.CellValue = new CellValue(data.Value.ToString());
            }
            else {
                SetCellEmptyValue(cell);
            }

        }
        internal void SetValue(Cell cell, double? data) {
            if (data.HasValue) {
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.CellValue = new CellValue(data.Value.ToString(CultureInfo.InvariantCulture));
            }
            else {
                SetCellEmptyValue(cell);
            }

        }
        internal void Save() {
            if (!disposed) {
                if (!(spreadsheetDocument is null)) {
                    spreadsheetDocument.Save();
                }
                else {
                }
            }
            else
                throw new ObjectDisposedException("Workbook");
        }
        internal static string GetColumnTitle(StringValue stringValue) {
            Throw.IfNull(stringValue);
            return Regex.Match(stringValue!.Value!, "[A-Z]+").Value;
        }
        private SpreadsheetDocument SpreadsheetDoc {
            get {
                if (spreadsheetDocument == null) {
                    spreadsheetDocument = SpreadsheetDocument.Open(file.FullName, isEditable);
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                    var sharedStringsPart = workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<SharedStringTablePart>();
                    sharedStrings = new SharedStrings(sharedStringsPart);
                    ;
                }
                return spreadsheetDocument;
            }
        }

        private static void SetCellEmptyValue(Cell cell) {
            // Create a new CellValue with an empty string
            var emptyValue = new CellValue(string.Empty);

            // Clear any existing cell text
            cell.RemoveAllChildren<Text>();

            // Add the empty value to the cell
            cell.AppendChild(emptyValue);
        }
        private static Hyperlinks GetHyperlinksPart(WorksheetPart worksheetPart) {
            Hyperlinks? hyperlinks = worksheetPart.Worksheet.Descendants<Hyperlinks>().FirstOrDefault();
            if (hyperlinks is null) {
                PageMargins pageMargins = worksheetPart.Worksheet.Descendants<PageMargins>().First();
                hyperlinks = new Hyperlinks();
                worksheetPart.Worksheet.InsertBefore(hyperlinks, pageMargins);
            }
            return hyperlinks;
        }
        internal void Dispose(bool disposing) {
            if (!disposedValue) {
                if (disposing) {
                    spreadsheetDocument?.Dispose();
                    stream?.Dispose();

                }

                // free unmanaged resources (unmanaged objects) and override finalizer
                sharedStrings = null;
                spreadsheetDocument = null;
                disposedValue = true;
                stream = null;
            }
        }

        public void Dispose() {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }

}
