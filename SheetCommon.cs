using Bs.XML.SpreadSheet.Resources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using SheetRow = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace Bs.XML.SpreadSheet {
    /// <summary>
    /// Представляет содержание значений ячеек листа Excel.
    /// </summary>
    public class SheetCommon : IEnumerable<SheetCommon.Row> {
        public class Writable : SheetCommon, IEnumerable<EditRow> {
            internal Writable(IEnumerable<string[]> rows, IEnumerable<string> titles) : base(rows, titles) {
            }
            IEnumerator<EditRow> IEnumerable<EditRow>.GetEnumerator() {
                for (int i = 0; i < rows.Count; i++) {
                    yield return new EditRow(rows[i], titles);
                }
            }
            public new EditRow this[int index] {
                get { return new EditRow(rows[index], titles); }
            }
            public string[] Add(string[] row) {
                string[] localRow = new string[columnsCount];
                Array.Copy(row, localRow, Math.Min(columnsCount, row.Length));
                rows.Add(localRow);
                return localRow;
            }
        }
        public readonly struct EditRow {
            //private readonly SheetCommon parent;
            private readonly string[] row;
            private readonly string[] titles;
            internal EditRow(string[] row, string[] titles) {
                this.row = row;
                this.titles = titles;
            }
            public string this[string columnTitle] {
                get {
                    int columnIndex = Array.IndexOf(titles, columnTitle);
                    return row[columnIndex];
                }
                set {
                    int columnIndex = Array.IndexOf(titles, columnTitle);
                    row[columnIndex] = value;
                }
            }
        }
        public struct Row : IReadOnlyDictionary<string, string> {
            public readonly int Index;
            private readonly SheetCommon parent;
            private KeyValuePair<string, string>[]? thisRow;
            internal Row(int rowIndex, SheetCommon parent) {
                thisRow = null;
                this.Index = rowIndex;
                this.parent = parent;
            }
            public string this[string columnTitle] {
                get {
                    return parent?[Index, columnTitle] ?? string.Empty;
                }
            }
            public bool TryGetValue(string title, out string value) {
                if (!string.IsNullOrEmpty(title)) {
                    int columnIndex = Array.IndexOf(parent.titles, title);
                    if (columnIndex >= 0) {
                        value = parent.rows[Index][columnIndex];
                        return true;
                    }
                }
                value = string.Empty;
                return false;
            }

            public bool ContainsKey(string key) {
                return Array.IndexOf(parent.titles, key) >= 0;
            }

            public IEnumerator<KeyValuePair<string, string>> GetEnumerator() {
                if (thisRow is null) {
                    thisRow = new KeyValuePair<string, string>[Titles.Length];
                    var values = parent.rows[Index];
                    for (int i = 0; i < thisRow.Length; i++) {
                        thisRow[i] = new KeyValuePair<string, string>(Titles[i], values[i]);
                    }
                }
                IEnumerator<KeyValuePair<string, string>> enumerator = (IEnumerator<KeyValuePair<string, string>>)thisRow.GetEnumerator();
                return enumerator;
            }

            IEnumerator IEnumerable.GetEnumerator() {
                return GetEnumerator();
            }

            public string[] Titles => parent?.Titles ?? Array.Empty<string>();

            public IEnumerable<string> Keys => Titles;

            public IEnumerable<string> Values => parent.rows[Index];

            public int Count => Titles.Length;
        }
        private readonly List<string[]> rows;
        private int columnsCount;
        private string[] titles;
        public SheetCommon() : this(Array.Empty<string[]>(), Array.Empty<string>()) {
        }
        public SheetCommon(int columnsCount) {
            rows = new List<string[]>();
            this.columnsCount = columnsCount;
            titles = new string[columnsCount];
            for (int i = 0; i < columnsCount; i++) {
                titles[i] = String.Empty;
            }
        }
        public SheetCommon(IEnumerable<string> titles) : this(Array.Empty<string[]>(), titles) {
        }
        protected SheetCommon(IEnumerable<string[]> rows, IEnumerable<string> titles) {
            this.rows = new List<string[]>(rows);
            this.titles = titles.ToArray();
            this.columnsCount = this.titles.Length;
        }

        public int RowsCount => rows.Count;
        public int ColumnsCount => columnsCount;
        public string[] Titles =>  titles.ToArray();

        public Row this[int index] {
            get => new Row(index, this);
        }
        public string this[int rowIndex, int columnIndex] {
            get {
                return rows[rowIndex][columnIndex];
            }
        }
        public string this[int rowIndex, string columnTitle] {
            get {
                Throw.IfIsTrue(rowIndex < 0 | rowIndex >= rows.Count,
                    () => new IndexOutOfRangeException());
                int columnIndex = Array.IndexOf(titles, columnTitle);
                Throw.IfIsTrue(columnIndex == -1,
                    () => new IndexOutOfRangeException(FormattableStringFactory.Create(ExceptionMessage.TitleNotFound, columnTitle).ToString()));
                return rows[rowIndex][columnIndex];
            }
        }
        public int ColumnIndex(string title) {
            return Array.IndexOf(titles, title);
        }
        public async Task LoadAsync(string fullFileName, string sheetName) {
            await Task.Run(() => LoadResource(fullFileName, sheetName))
               .ConfigureAwait(false);
        }
        /// <summary>
        /// Загружает страницу xlsx как двумерную таблицу.
        /// </summary>
        /// <param name="fullFileName">Имя файла, включая полный или относительный путь.</param>
        /// <param name="sheetName">Имя листа в файле.</param>
        /// <exception cref="System.IO.FileLoadException"></exception>
        /// <exception cref="FormatException"></exception>
        public void LoadResource(string fullFileName, string sheetName) {
            StringBuilder error = new StringBuilder();

            using (Workbook book = new Workbook(fullFileName)) {

                SheetData? jobsSheet = book.GetWorksheetPart(sheetName).Worksheet.GetFirstChild<SheetData>()
                    ?? throw new System.IO.FileLoadException(FormattableStringFactory.Create(ExceptionMessage.MissingSheetInFile,fullFileName,sheetName).ToString());
                var sheetRowsCursor = book.GetRows(jobsSheet).GetEnumerator();
                List<KeyValuePair<string, string>> knownJobStepColumns = new List<KeyValuePair<string, string>>();
                try {
                    if (!sheetRowsCursor.MoveNext()) {
                        error.AppendLine(FormattableStringFactory.Create(ExceptionMessage.TheSheetMustContainThreeLines, sheetName).ToString());
                    }
                    SheetRow titlesRow = sheetRowsCursor.Current;
                    foreach (Cell cell in titlesRow.Elements<Cell>()) {
                        string? title = book.GetStringValue(cell)?.Trim();
                        if (string.IsNullOrWhiteSpace(title))
                            continue;
                        else if (knownJobStepColumns.Any(kvp => kvp.Value == title)) {
                            error.AppendLine(FormattableStringFactory.Create(ExceptionMessage.TheHeadingIsDuplicated, title).ToString());
                            continue;
                        }
                        string columnTitle = Workbook.GetColumnTitle(cell.CellReference!);
                        knownJobStepColumns.Add(new KeyValuePair<string, string>(columnTitle, title ?? string.Empty));
                    }
                    if (!sheetRowsCursor.MoveNext() || !sheetRowsCursor.MoveNext()) {
                        error.AppendLine(FormattableStringFactory.Create(ExceptionMessage.TheSheetMustContainThreeLines, sheetName).ToString());
                    }
                }
                catch (Exception ex) {
                    error.AppendLine(ex.ToString());
                }
                if (error.Length > 0)
                    throw new FormatException(error.ToString());
                columnsCount = knownJobStepColumns.Count;
                this.titles = knownJobStepColumns.Select(kvp => kvp.Value).ToArray();

                do {
                    SheetRow row = sheetRowsCursor.Current;
                    string[] thisRow = new string[columnsCount];
                    bool isRowHaveValues = false;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        string columnTitle = Workbook.GetColumnTitle(cell!.CellReference ?? string.Empty);
                        int i_col = knownJobStepColumns.FindIndex(kvp => kvp.Key == columnTitle);
                        if (i_col >= 0) {
                            var cellvalue = book.GetStringValue(cell).Trim();
                            if (string.IsNullOrEmpty(cellvalue))
                                thisRow[i_col] = string.Empty;
                            else {
                                isRowHaveValues = true;
                                thisRow[i_col] = cellvalue;
                            }
                        }
                    }
                    if (isRowHaveValues) { rows.Add(thisRow); }
                } while (sheetRowsCursor.MoveNext());
            }
        }
        /// <summary>
        /// Сохраняет таблицу в шаблон ресурса.
        /// </summary>
        /// <param name="fullFileName"></param>
        /// <param name="sheetName"></param>
        /// <remarks>В шаблоне из первой строки считываются заголовки столбцов. 
        /// Вторая строка шаблона является строкой комментариев и не используется.
        /// Заполнение таблицы начинается с третьей строки; вносятся столбцы ресурса,
        /// заголовки которых найдены в первой строке.</remarks>
        /// <exception cref="System.IO.FileLoadException"></exception>
        /// <exception cref="FormatException"></exception>
        public void SaveToResource(string templateName, string fullFileName, string sheetName) {
            var fiTemplate = new FileInfo(templateName);
            var fiTarget = new FileInfo(fullFileName);
            if (!fiTemplate.Exists)
                throw new FileLoadException(ExceptionMessage.TemplateFileNotFound);

            using (Workbook book = Workbook.FromTemplate(fiTemplate.FullName, fiTarget.FullName)) {
                DocumentFormat.OpenXml.Packaging.WorksheetPart wsPart = book.GetWorksheetPart(sheetName);
                SheetData jobsSheet = wsPart.Worksheet.GetFirstChild<SheetData>()
                    ?? throw new System.IO.FileLoadException(FormattableStringFactory.Create(ExceptionMessage.MissingSheetInFile, fullFileName, sheetName).ToString());
                List<KeyValuePair<string, string>> templateColumns = new List<KeyValuePair<string, string>>(); // key title=>value columnTite
                {
                    SheetRow row = book.Row(wsPart, 1)
                    ?? throw new FormatException(FormattableStringFactory.Create(ExceptionMessage.TheSheetMustContainHeaderRow, sheetName).ToString());
                    // Чтение строки заголовков
                    //SheetRow titlesRow = sheetRowsCursor.Current;
                    StringBuilder error = new StringBuilder();
                    foreach (Cell cell in row.Elements<Cell>()) {
                        string? title = book.GetStringValue(cell)?.Trim();
                        if (string.IsNullOrWhiteSpace(title))
                            continue;
                        else if (templateColumns.Any(kvp => kvp.Key == title)) {
                            error.AppendLine(FormattableStringFactory.Create(ExceptionMessage.TheHeadingIsDuplicated, title).ToString());
                            continue;
                        }
                        else if (!Titles.Contains(title)) { continue; }
                        string columnTitle = (cell.CellReference is null) ? string.Empty : Workbook.GetColumnTitle(cell.CellReference);
                        templateColumns.Add(new KeyValuePair<string, string>(title!, columnTitle));
                    }
                    if (error.Length > 0)
                        throw new FormatException(error.ToString());

                    if (templateColumns.Count == 0) { // Нет совпадающих столбцов, так что ничего заполнять и не потребуется.
                        book.Save();
                        return;
                    }
                }

                foreach (Row item in this) {
                    UInt32Value rowIndex = (uint)item.Index + 3u;
                    SheetRow sheetRow = book.GetOrAddRow(wsPart, rowIndex);
                    sheetRow.RemoveAllChildren();
                    for (int i = 0; i < templateColumns.Count; i++) {
                        string column = templateColumns[i].Value;
                        string title = templateColumns[i].Key;
                        string value = item[title];
                        var cell = new Cell() { CellReference = $"{column}{(uint)rowIndex}" };
                        book.SetValue(cell, value);
                        sheetRow.Append(cell);
                    }
                }
                wsPart.Worksheet.Save();
                book.Save();
            }
        }

        /// <summary>
        /// Создает копию экземпляра SheetCommon, с возможностью прямой записи элементов таблицы.
        /// </summary>
        /// <returns></returns>
        public Writable GetWritable() {
            return new Writable(rows, titles);

        }
        public IEnumerator<SheetCommon.Row> GetEnumerator() {
            for (int i = 0; i < RowsCount; i++) {
                yield return this[i];
            }
        }
        IEnumerator IEnumerable.GetEnumerator() {
            return GetEnumerator();
        }

        internal Func<Row, string> GetterFromRow(string title) {
            int columnIndex = ColumnIndex(title);
            if (columnIndex >= 0)
                return row => GetValue(row, columnIndex);
            throw new IndexOutOfRangeException(FormattableStringFactory.Create(ExceptionMessage.TitleNotFound, title).ToString());
            ;
        }
        internal Func<string[], string> GetterFromArray(string title) {
            int columnIndex = ColumnIndex(title);
            if (columnIndex >= 0)
                return row => GetValue(row, columnIndex);
            throw new IndexOutOfRangeException(FormattableStringFactory.Create(ExceptionMessage.TitleNotFound, title).ToString());
        }
        private string GetValue(string[] row, int columnIndex) {
            return row[columnIndex];
        }
        private string GetValue(SheetCommon.Row row, int columnIndex) {
            return rows[row.Index][columnIndex];
        }

        public IEnumerable<string[]> GetInternalRows() {
            return rows;
        }
    }
}
