using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using keypsDownloaderCore.Models;
using NodaTime;
using NodaTime.Text;

namespace keypsDownloaderCore {
    public class KeypsWorksheetParser {
        private string SpreadsheetName { get; set; }

        public KeypsWorksheetParser(string name) {
            if (!File.Exists(name)) {
                throw new FileNotFoundException($"{name} does not exists.");
            }

            SpreadsheetName = name;
        }

        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id) {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        String GetCellValueAsString(WorkbookPart wbPart, Cell cell) {
            string cellValue = string.Empty;

            if (cell.DataType != null) {
                if (cell.DataType == CellValues.InlineString) cellValue = cell.InnerText;
                else if (cell.DataType == CellValues.SharedString) {
                    int id = -1;

                    if (Int32.TryParse(cell.InnerText, out id)) {
                        SharedStringItem item = GetSharedStringItemById(wbPart, id);

                        if (item.Text != null) {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null) {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null) {
                            cellValue = item.InnerXml;
                        }
                    }
                }
                else if (cell.DataType != CellValues.InlineString)
                    cellValue = cell.CellValue.Text;

                return cellValue;
            }

            if (cell.InnerText != null) return cell.InnerText;
            return String.Empty;
        }

        Cell GetCellByReference(WorkbookPart wbPart, string cellReference) {
            WorksheetPart wsPart = wbPart.WorksheetParts.FirstOrDefault();
            Cell cell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellReference);
            return cell;
        }

        string GetCellValueByReference(WorkbookPart wbPart, string cellReference) {
            return GetCellValueAsString(wbPart, GetCellByReference(wbPart, cellReference));
        }

        ColumnInfo BuildColumnInfo(WorkbookPart wbPart, WorksheetPart wsPart, SheetData sheetData) {
            ColumnInfo colInfo = new ColumnInfo();

            foreach (Row row in sheetData.Elements<Row>()) {
                foreach (Cell cell in row.Elements<Cell>()) {
                    if (!colInfo.IsCompleted()) {
                        String cellValue = GetCellValueAsString(wbPart, cell);
                        Regex regex = new Regex("[A-Za-z]+");
                        Match columnHeader = regex.Match(cell.CellReference);
                        Match rowIndex = Regex.Match(cell.CellReference, "[1-9]+");

                        if (cellValue == "ID") colInfo.IdColumn = columnHeader.Value;
                        else if (cellValue == "Tarih") colInfo.DateColumn = columnHeader.Value;
                        else if (cellValue == "Saat") colInfo.TimeColumn = columnHeader.Value;
                        else if (cellValue == "SINIF") colInfo.GradeColumn = columnHeader.Value;
                        else if (cellValue == "Anabilimdalı") colInfo.AlanColumn = columnHeader.Value;
                        else if (cellValue == "Konu") colInfo.KonuColumn = columnHeader.Value;
                        else if (cellValue == "Eğitici") colInfo.TeacherColumn = columnHeader.Value;
                        else if (cellValue == "#") {
                            int
                                count = 1; // Default +1 because we want an example of the value of column not the column header
                            string nextVal = GetCellValueByReference(wbPart,
                                $"{columnHeader.Value}{Int32.Parse(rowIndex.Value) + count}");

                            while (nextVal == "-") {
                                count += 1;
                                nextVal = GetCellValueByReference(wbPart,
                                    $"{columnHeader.Value}{Int32.Parse(rowIndex.Value) + count}");
                            }

                            if (nextVal.Length == 54) colInfo.MeetingIdColumn = columnHeader.Value;
                            else {
                                if (nextVal.Contains("kapitta")) {
                                    colInfo.KapittaColumn = columnHeader.Value;
                                }
                            }
                        }
                    }
                }
            }

            return colInfo;
        }

        Grade GetGradeFromString(string gradeString) {
            if (gradeString == "1. Sınıf") return Grade.First;
            else if (gradeString == "2. Sınıf") return Grade.Second;
            else if (gradeString == "3. Sınıf") return Grade.Third;
            else if (gradeString == "4. Sınıf") return Grade.Fourth;
            else if (gradeString == "5. Sınıf") return Grade.Fifth;
            else if (gradeString == "6. Sınıf") return Grade.Sixth;
            else return Grade.Undefined;
        }

        public async Task<List<Lesson>> ParseAsync() {
            return await Task.Run(Parse);
        }

        public List<Lesson> Parse() {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(SpreadsheetName, false)) {
                WorkbookPart wbPart = doc.WorkbookPart;
                WorksheetPart wsPart = wbPart.WorksheetParts.First();
                SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();

                ColumnInfo colInfo = BuildColumnInfo(wbPart, wsPart, sheetData);

                List<Lesson> lessonList = new List<Lesson>();

                LocalDateTimePattern datePattern =
                    LocalDateTimePattern.Create("%dd/%MM/%yyyy %HH:%mm", CultureInfo.CurrentCulture);

                foreach (Row row in sheetData.Elements<Row>()) {
                    try {
                        string alanColVal = GetCellValueByReference(wbPart, $"{colInfo.AlanColumn}{row.RowIndex}");
                        if (alanColVal != "Anabilimdalı") {
                            string baseKapittaUrl = Regex.Match(
                                GetCellValueByReference(wbPart, $"{colInfo.KapittaColumn}{row.RowIndex}"),
                                @"(?<base_url>.+)\/bigbluebutton\/api").Groups["base_url"].Value;

                            string meetingId =
                                GetCellValueByReference(wbPart, $"{colInfo.MeetingIdColumn}{row.RowIndex}");
                            string url =
                                $"{baseKapittaUrl}/playback/presentation/2.0/playback.html?meetingId={meetingId}";
                            string konu = GetCellValueByReference(wbPart, $"{colInfo.KonuColumn}{row.RowIndex}");
                            string alan = GetCellValueByReference(wbPart, $"{colInfo.AlanColumn}{row.RowIndex}");
                            string idString = GetCellValueByReference(wbPart, $"{colInfo.IdColumn}{row.RowIndex}");
                            int id = Int32.Parse(idString);
                            string teacher = GetCellValueByReference(wbPart, $"{colInfo.TeacherColumn}{row.RowIndex}");

                            string dateString = GetCellValueByReference(wbPart, $"{colInfo.DateColumn}{row.RowIndex}");
                            string timeString = GetCellValueByReference(wbPart, $"{colInfo.TimeColumn}{row.RowIndex}");
                            string dateTimeString = $"{dateString} {timeString}";
                            LocalDateTime date = datePattern.Parse(dateTimeString).Value;

                            string gradeString =
                                GetCellValueByReference(wbPart, $"{colInfo.GradeColumn}{row.RowIndex}");
                            Grade grade = GetGradeFromString(gradeString);

                            lessonList.Add(new Lesson(id, konu, alan, url, baseKapittaUrl, date, grade, teacher));
                        }
                    }
                    catch (NullReferenceException e) {
                        Debug.WriteLine($"An error happened {e}");
                    }
                }

                var duplicateKeys = lessonList.
                    GroupBy(s => s.Name).
                    Where(g => g.Count() > 1).
                    Select(x => x.Key);

                foreach (var key in duplicateKeys) {
                    var duplicates = lessonList.Where(x => x.Name == key).OrderBy(x => x.Date);
                    foreach (var (item, index) in duplicates.Select((val, i) => (val, i))) {
                        item.Name = $"{item.Name} - {index + 1}";
                    }
                }

                return lessonList;
            }
        }
    }
}