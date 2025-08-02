using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Package
{
    internal class StylePartHelper
    {
        public static void AddStylesPart(SpreadsheetDocument spreadsheet)
        {
            var stylesPart = spreadsheet.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            var numberingFormats = new NumberingFormats();
            numberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 164U,
                FormatCode = "yyyy/\\ mm/\\ dd\\.\\ hh:mm;@"
            });
            numberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = 165U,
                FormatCode = "_-* #,##0\\ \"Ft\"_-;\\-* #,##0\\ \"Ft\"_-;_-* \"-\"??\\ \"Ft\"_-;_-@_-"

            });

            var fonts = new Fonts() { Count = 1, KnownFonts = true };
            fonts.AppendChild(new Font(
                new FontSize { Val = 11 },
                new Color { Rgb = "000000" },
                new FontName { Val = "Calibri" },
                new FontFamilyNumbering { Val = 2 },
                new FontScheme { Val = FontSchemeValues.Minor }
            ));

            var fills = new Fills() { Count = 1 };
            fills.AppendChild(new Fill(
                new PatternFill { PatternType = PatternValues.None }
            ));

            var borders = new Borders() { Count = 1 };
            borders.AppendChild(new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder()
            ));

            stylesPart.Stylesheet.AppendChild(numberingFormats);
            stylesPart.Stylesheet.AppendChild(fonts);
            stylesPart.Stylesheet.AppendChild(fills);
            stylesPart.Stylesheet.AppendChild(borders);
            stylesPart.Stylesheet.AppendChild(new CellStyleFormats(
                new CellFormat
                {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U
                }
            ));
            stylesPart.Stylesheet.AppendChild(new CellFormats(
                new CellFormat
                {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U,
                    ApplyNumberFormat = true
                },
                new CellFormat
                {
                    NumberFormatId = 164U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U,
                    ApplyNumberFormat = true
                },
                new CellFormat
                {
                    NumberFormatId = 165U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U,
                    ApplyNumberFormat = true
                }
            ));
        }
    }
}
