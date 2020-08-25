using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public class ArchivoExcelBase
    {
        private const double TAMANIO_LETRA = 0.96;
        private const double TAMANIO_CELDA_POR_DEFECTO = 12.50;
        private SpreadsheetDocument _documentoExcel { get; set; }
        private WorkbookPart _libro { get; set; }
        private WorksheetPart _hojaCalculo { get; set; }
        private Sheets _hojas { get; set; }
        private SheetData _sheetData { get; set; }
        private UInt32 _numeroFila { get; set; }
        private string[] _letras { get; set; }        
        private int _indiceLetra { get; set; }
        private string _direccion { get; set; }     
        protected string _rutaArchivo { get; set; }
        private DataTable _fuenteDeDatos;
        public DataTable FuenteDeDatos
        {
            get { return _fuenteDeDatos; }
            set 
            {
                if (FuenteDeDatos.Rows.Count > 26)
                    _letras = ObtenerLetras(FuenteDeDatos.Rows.Count);

                _fuenteDeDatos = value; 
            }
        }
        
        private Worksheet _workSheet;
        public Worksheet WorkSheet
        {
            get 
            { 
                _workSheet = _hojaCalculo.Worksheet; 
                return _workSheet;
            }
            set { _workSheet = value; }
        }
      
        public string Titulo { get; set; }
        public string NombreHoja { get; set; }    
        public string[] Encabezados { get; set; }
        public delegate void AgregarEncabezadosHandler();
        public delegate void AgregarInformacionHandler();
        public delegate void EstablecerAnchoColumnasHandler();
        public AgregarEncabezadosHandler EncabezadosHandler { get; set; }
        public AgregarInformacionHandler InformacionHandler { get; set; }
        public EstablecerAnchoColumnasHandler AnchoColumnasHandler { get; set; }


        public ArchivoExcelBase(string rutaArchivo, string titulo, string nombreHoja)
        {
            this._rutaArchivo = rutaArchivo;
            this.Titulo = titulo;
            this.NombreHoja = nombreHoja;

            Inicializar();
        }

        private void Inicializar()
        {
            _letras = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            _indiceLetra = 0;
            _direccion = "A1";
            _numeroFila = UInt32.Parse("1");
        }

        protected void CrearLibro()
        {
            _documentoExcel = SpreadsheetDocument.Create(this._rutaArchivo, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = _documentoExcel.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            workbookpart.Workbook.Save();

            _libro = workbookpart;
            _hojas = _documentoExcel.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
        }

        protected void CrearHoja(string nombreLibro, UInt32 index)
        {
            WorksheetPart worksheetPart = _libro.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            _hojaCalculo = worksheetPart;
            AdicionarHoja(nombreLibro, index);

            _sheetData = _hojaCalculo.Worksheet.GetFirstChild<SheetData>();
        }

        protected virtual void AdicionarHoja(string nombreLibro, UInt32 indexTag)
        {
            Sheet hoja = new Sheet()
            {
                Id = this._documentoExcel.WorkbookPart.GetIdOfPart(_hojaCalculo),
                SheetId = indexTag,
                Name = nombreLibro
            };

            _hojas.Append(hoja);
        }

        public void Guardar()
        {
            if (_libro != null)
                _libro.Workbook.Save();
        }

        public void CerrarLibro()
        {
            if (_documentoExcel != null)
                _documentoExcel.Close();
        }

        protected string[] ObtenerLetras(int cantidadLetras)
        {
            int longuitud = 26;
		    int iteraciones = (cantidadLetras / longuitud) - 1;

            return _letras
                .Select((letra, indice) => new { letra, indice })
                .Where(o => o.indice < iteraciones)
                .Prepend(new { letra = "", indice = -1 })
                .Select(o => new { letra = o.letra, letras = _letras })
                .SelectMany(o => o.letras, (letra1, letra2) => letra1.letra + letra2)
                .ToArray(); 
        }

        protected virtual void CalcularAnchoCelda()
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match;

            var celdas = _sheetData.Descendants<Cell>();

            var groupby = celdas
                .GroupBy(c =>
                {
                    match = regex.Match(c.CellReference.ToString());
                    return match.Value;
                })
                .Select((g, index) => new { Letra = g.Key, Puntos = g.Max(cv => cv.CellValue.InnerText.Length) * TAMANIO_LETRA })
                .OrderBy(o => o.Letra)
                .Select((o, i) => new { Indice = i + 1, o.Puntos })
                .Where(o => o.Puntos > TAMANIO_CELDA_POR_DEFECTO);

            foreach (var item in groupby)
            {
                EstablecerAnchoCelda(item.Indice, item.Puntos);
            }
        }

        public void EstablecerAnchoCelda(int indice, DoubleValue ancho)
        {
            uint Index = (uint)indice;

            Columns cs = _workSheet.GetFirstChild<Columns>();
            if (cs != null)
            {
                IEnumerable<Column> ic = cs.Elements<Column>().Where(r => r.Min == Index).Where(r => r.Max == Index);
                if (ic.Count() > 0)
                {
                    Column c = ic.First();
                    c.Width = ancho;
                }
                else
                {
                    Column c = new Column() { Min = Index, Max = Index, Width = ancho, CustomWidth = true };
                    cs.Append(c);
                }
            }
            else
            {
                cs = new Columns();
                Column c = new Column() { Min = Index, Max = Index, Width = ancho, CustomWidth = true };
                cs.Append(c);
                _workSheet.InsertAfter(cs, _workSheet.GetFirstChild<SheetFormatProperties>());
            }
        }

        public void CombinarCeldas(string celdaInicial, string celdaFinFinal)
        {
            MergeCells mergeCells;

            if (_workSheet.Elements<MergeCells>().Count() > 0)
                mergeCells = _workSheet.Elements<MergeCells>().First();
            else
            {
                mergeCells = new MergeCells();

                if (_workSheet.Elements<CustomSheetView>().Count() > 0)
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<CustomSheetView>().First());
                else
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<SheetData>().First());
            }
            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(celdaInicial + ":" + celdaFinFinal) };
            mergeCells.Append(mergeCell);
        }
    }
}