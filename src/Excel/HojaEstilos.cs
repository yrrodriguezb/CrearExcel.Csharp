using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public class HojaEstilos
    {
        public static UInt32Value TITULO = 1;
        public static UInt32Value SUBTITULO = 2;
        public static UInt32Value ENCABEZADO_TABLA = 3;
        public static UInt32Value DATO_NORMAL = 4;
        public static UInt32Value DATO_NUMERICO = 5;
        public static UInt32Value DATO_MONEDA = 6;
        public static UInt32Value DATO_FECHA = 7;


        public static Stylesheet GenerarEstilos()
        {
            Stylesheet Estilos =
            new Stylesheet(
                new Fonts(
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 8 }),
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 10 }, new Bold { Val = true }, new Color { Rgb = new HexBinaryValue() { Value = "007bff" } }),
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 10 }, new Bold { Val = true }, new Color { Rgb = new HexBinaryValue() { Value = "17a2b8" } }),
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 8 }, new Bold { Val = true }, new Color { Rgb = new HexBinaryValue() { Value = "000000" } })
                    ),
                new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "c3e6cb" } }) { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
                    new Border(
                        new LeftBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    new CellFormat() { FillId = 0, BorderId = 0 },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 1, FillId = 0, BorderId = 0 },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 2, FillId = 0, BorderId = 0 },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 3, FillId = 2, BorderId = 1 },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FillId = 0, BorderId = 1 },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right }) { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 1, ApplyNumberFormat = true, ApplyAlignment = true, ApplyFont = true, ApplyFill = true },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right }) { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 39, ApplyNumberFormat = true, ApplyAlignment = true, ApplyFont = true, ApplyFill = true },
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right }) { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 14, ApplyNumberFormat = true, ApplyAlignment = true, ApplyFont = true, ApplyFill = true }
                )
            );
            return Estilos;
        }
    }
}