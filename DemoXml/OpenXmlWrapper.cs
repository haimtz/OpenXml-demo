﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DemoXml
{
    public class OpenXmlWrapper
    {
        /// <summary>
        /// Get the spread document
        /// </summary>
        /// <param name="filename">file name of the document</param>
        /// <returns>instance of file</returns>
        public SpreadsheetDocument Document(string filename)
        {
            SpreadsheetDocument document = null;

            try
            {
                // Create document
                document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook, false);
                document.AddWorkbookPart();

                var workpart = document.WorkbookPart;
                workpart.Workbook = new Workbook();
                workpart.Workbook.Save();

                var sharedStringTablePart = workpart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
                sharedStringTablePart.SharedStringTable.Save();

                // Create sheets
                workpart.Workbook.Sheets = new Sheets();
                workpart.Workbook.Save();

                var styles = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = CreatStylesheet();
                styles.Stylesheet.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return document;
        }

        public Worksheet AddSheet(WorkbookPart workbookPart, string name)
        {
            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            // add single Sheet
            var workSheetpart = workbookPart.AddNewPart<WorksheetPart>();
            workSheetpart.Worksheet = new Worksheet(new SheetData());
            workSheetpart.Worksheet.InsertAt(new Columns(), 0);
            workSheetpart.Worksheet.Save();

            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(workSheetpart),
                SheetId = (uint)(workbookPart.Workbook.Sheets.Count() + 1),
                Name = name
            };

            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return workSheetpart.Worksheet;
        }

        public void AddRow(Worksheet worksheet, bool isHeader, params object[] values)
        {
            var row = new Row();
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var styleIndex = StyleConst.CellFormat.CellFormatStyle.CELL_REGULAR;

            if (isHeader)
                styleIndex = StyleConst.CellFormat.CellFormatStyle.CELL_HEADER;
            
            foreach (var value in values)
            {
                var cell = GetCell(value, styleIndex);
                row.AppendChild(cell);
            }

            sheetData.Append(row);
            worksheet.Save();
        }

        public void SetWidth(Worksheet worksheet, int columnIndex, int width)
        {
            //worksheet.InsertAt(new Columns(), 0);
            var columns = worksheet.Elements<Columns>().ElementAt(0);

            var col1 = new Column
            {
                Min = 1,
                Max = 1,
                CustomWidth = true,
                Width = 10,
            };

            var col2 = new Column
            {
                Min = 2,
                Max = 2,
                Width = 50,
            };

            columns.Append(col1, col2);
            worksheet.Save();
        }

        private Stylesheet CreatStylesheet()
        {
            var style = new Stylesheet();

            #region Font Style
            style.InsertAt(new Fonts(), StyleConst.FontsConst.FONT);
            style.GetFirstChild<Fonts>()
                .InsertAt<Font>(
                    new Font
                    {
                        FontSize = new FontSize {Val = 11},
                        Bold = new Bold {Val = true}
                    }, StyleConst.FontsConst.FontStyle.FONT_BOLD);

                style.GetFirstChild<Fonts>().InsertAt<Font>(new Font
                {
                    FontSize = new FontSize {Val = 11},
                    Bold = new Bold {Val = false}
                }, StyleConst.FontsConst.FontStyle.FONT_REGULAR);
            #endregion

            #region Fill Style
            style.InsertAt(new Fills(), StyleConst.FillConst.FILL);
            style.GetFirstChild<Fills>().InsertAt<Fill>(
               new Fill
               {
                   PatternFill = new PatternFill
                   {
                       PatternType = new EnumValue<PatternValues>
                       {
                           Value = PatternValues.Gray125
                       },
                       ForegroundColor = new ForegroundColor
                       {
                           Rgb = "FFFF00"
                       },
                       BackgroundColor = new BackgroundColor
                       {
                           Indexed = (UInt32Value)64U,
                           Rgb = "FFFF00"
                       }
                   }
               }, StyleConst.FillConst.FillStyle.REGULAR);

            #endregion

            #region Border Style
            style.InsertAt(new Borders(), 2/*StyleConst.BorderConst.BORDER*/);
            style.GetFirstChild<Borders>().InsertAt<Border>(
               new Border
               {
                   LeftBorder = new LeftBorder() { Style = BorderStyleValues.Thick },
                   RightBorder = new RightBorder { Style = BorderStyleValues.Thick },
                   TopBorder = new TopBorder { Style = BorderStyleValues.Thick },
                   BottomBorder = new BottomBorder { Style = BorderStyleValues.Thick },
                   DiagonalBorder = new DiagonalBorder()
               }, StyleConst.BorderConst.Bordertyle.REGULAR);
            #endregion

            #region Cell Format
            style.InsertAt(new CellFormats(), StyleConst.CellFormat.CELL_FORMAT_STYLE);
            style.GetFirstChild<CellFormats>().InsertAt<CellFormat>(
                new CellFormat
                {
                    FontId = StyleConst.FontsConst.FontStyle.FONT_BOLD,
                    NumberFormatId = 0,
                    FillId = 0,
                    BorderId = 0
                }, StyleConst.CellFormat.CellFormatStyle.CELL_HEADER);

            style.GetFirstChild<CellFormats>().InsertAt<CellFormat>(
                new CellFormat
                {
                    FontId = StyleConst.FontsConst.FontStyle.FONT_REGULAR,
                    NumberFormatId = 0,
                    FillId = 0,
                    BorderId = 0
                }, StyleConst.CellFormat.CellFormatStyle.CELL_REGULAR);

            
            #endregion

            return style;
        }


        private Cell GetCell(object value, int styleIndex)
        {
            var type = value.GetType().Name == "String" ? CellValues.String : CellValues.Number;

            var cell = new Cell
            {
                DataType = type,
                CellValue = new CellValue(value.ToString()),
                StyleIndex = (UInt32)styleIndex,
                
            };

            return cell;
        }

    }
}
