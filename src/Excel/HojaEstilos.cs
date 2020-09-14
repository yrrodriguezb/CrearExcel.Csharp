using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public class HojaEstilos
    {
        private static string COLOR_ENCABEZADO_TABLA = "cfd8dc";
        private static string COLOR_BORDE = "eceff1";
        private static string COLOR_TITULO = "263238";
        private static string COLOR_SUBTITULO = "455a64";
        private static string COLOR_NIVEL_UNO = "CFDEC7";
        private static string COLOR_NIVEL_DOS = "C2D6CD";
        private static string COLOR_NIVEL_TRES = "B6CDC3";
        public static string COLOR_NIVEL_CUATRO = "AAC5C9";
        private static string COLOR_NIVEL_CINCO = "9EBDAE";
        private static string COLOR_NIVEL_SEIS = "92b5a4";

        public static UInt32Value TITULO = 1;
        public static UInt32Value SUBTITULO = 2;
        public static UInt32Value ENCABEZADO_TABLA = 3;
        public static UInt32Value DATO_NORMAL = 4;
        public static UInt32Value DATO_NUMERICO = 5;
        public static UInt32Value DATO_MONEDA = 6;
        public static UInt32Value DATO_FECHA = 7;
        public static UInt32Value NIVEL_UNO = 8;
        public static UInt32Value NIVEL_DOS = 9;
        public static UInt32Value NIVEL_TRES = 10;
        public static UInt32Value NIVEL_CUATRO = 11;
        public static UInt32Value NIVEL_CINCO = 12;
        public static UInt32Value NIVEL_SEIS = 13;



        public static Stylesheet GenerarEstilos()
        {
            Stylesheet Estilos =
            new Stylesheet(
                new Fonts(
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 8 }),
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 10 }, new Bold { Val = true }, new Color { Rgb = new HexBinaryValue() { Value = COLOR_TITULO} }),
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 10 }, new Bold { Val = true }, new Color { Rgb = new HexBinaryValue() { Value = COLOR_SUBTITULO } }),
                        new Font(new FontName() { Val = "Verdana" }, new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 8 }, new Bold { Val = true }, new Color { Rgb = new HexBinaryValue() { Value = COLOR_TITULO } })
                    ),
                new Fills(
                /*0*/    new Fill(new PatternFill() { PatternType = PatternValues.None }),
                /*1*/    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
                /*2*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_ENCABEZADO_TABLA } }) { PatternType = PatternValues.Solid }),
                /*3*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_NIVEL_UNO } }) { PatternType = PatternValues.Solid }),
                /*4*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_NIVEL_DOS } }) { PatternType = PatternValues.Solid }),
                /*5*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_NIVEL_TRES } }) { PatternType = PatternValues.Solid }),
                /*6*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_NIVEL_CUATRO } }) { PatternType = PatternValues.Solid }),
                /*7*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_NIVEL_CINCO } }) { PatternType = PatternValues.Solid }),
                /*8*/    new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = COLOR_NIVEL_SEIS } }) { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
                    new Border(
                        new LeftBorder(new Color() { Rgb = COLOR_BORDE }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Rgb = COLOR_BORDE }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Rgb = COLOR_BORDE }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Rgb = COLOR_BORDE }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    /*0*/ new CellFormat() { FillId = 0, BorderId = 0 },
                    /*1*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 1, FillId = 0, BorderId = 0 },
                    /*2*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 2, FillId = 0, BorderId = 0 },
                    /*3*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 3, FillId = 2, BorderId = 1 },
                    /*4*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FillId = 0, BorderId = 1 },
                    /*5*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right }) { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 1, ApplyNumberFormat = true, ApplyAlignment = true,  ApplyFont = true, ApplyFill = true },
                    /*6*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 39, ApplyNumberFormat = true, ApplyAlignment = true,  ApplyFont = true, ApplyFill = true },
                    /*7*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 14, ApplyNumberFormat = true, ApplyAlignment = true, ApplyFont = true, ApplyFill = true },
                    /*8*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 3, FillId = 3, BorderId = 1 },
                    /*9*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 3, FillId = 4, BorderId = 1 },
                    /*10*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 3, FillId = 5, BorderId = 1 },
                    /*11*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 3, FillId = 6, BorderId = 1 },
                    /*12*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 3, FillId = 7, BorderId = 1 },
                    /*12*/ new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left }) { FontId = 3, FillId = 8, BorderId = 1 }
                )
            );
            return Estilos;
        }
    }
}