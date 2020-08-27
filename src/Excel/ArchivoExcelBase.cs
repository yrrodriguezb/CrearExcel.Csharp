using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public delegate void AgregarSubtitulosHandler();
    public delegate void AgregarEncabezadosHandler();
    public delegate void AgregarInformacionHandler();
    public delegate void EstablecerAnchoColumnasHandler();


    public class ArchivoExcelBase
    {
        private string[] _arrLetras = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        private const double TAMANIO_LETRA = 1.0;
        private const double TAMANIO_CELDA_POR_DEFECTO = 12.15;
        private SpreadsheetDocument _documentoExcel { get; set; }
        private WorkbookPart _libro { get; set; }
        private WorksheetPart _hojaCalculo { get; set; }
        private Sheets _hojas { get; set; }
        private SheetData _sheetData { get; set; }
        private UInt32 _numeroFila { get; set; }     
        private int _indiceLetra { get; set; }  
        protected string _rutaArchivo { get; set; }

        private string[] _letras;
        public string[] Letras
        {
            get { return _letras; }
            private set { _letras = value; }
        }

        private DataTable _fuenteDeDatos;
        public DataTable FuenteDeDatos
        {
            get { return _fuenteDeDatos; }
            set { _fuenteDeDatos = value; CargarLetras(); }
        }

        private string[] _encabezados; 
        public string[] Encabezados 
        { 
            get { return _encabezados; } 
            set { _encabezados = value; CargarLetras(); }
        }

        private int _longuitudColumnas;
        public int LonguitudColumnas
        {
            get
            {
                _longuitudColumnas = GetLonguitudColumnas(); 
                return _longuitudColumnas; 
            }
            set { _longuitudColumnas = value; }
        }
        
        private Worksheet _workSheet;
        public Worksheet WorkSheet
        {
            get { return _workSheet; }
            private set { _workSheet = value; }
        }
      
        public string Titulo { get; set; }
        public string NombreHoja { get; set; }
        public string[] ExcluirColumnas { get; set; }
        public AgregarSubtitulosHandler SubtitulosHandler { get; set; }
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

        public Row GetFila()
        {
            Row fila = new Row { RowIndex = _numeroFila };
            _numeroFila++;
            return fila;
        }

        public UInt32 GetNumeroFilaActual()
        {
            var numeroFila = _numeroFila;

            if (numeroFila > 1)
                return numeroFila - 1;

            return numeroFila;
        }

        public string GetLetra(int indice)
        {
            if (indice < 0 || indice > _letras.Count())
                throw new ArgumentOutOfRangeException(nameof(_letras));

            return _letras[indice].ToString();
        }

        public void SetIndiceLetra(int valor)
        {
            if (valor < 0)
                throw new ArgumentException("No se puede asignar un indice negativo", nameof(_indiceLetra));

            _indiceLetra = valor;
        }

        private void Inicializar()
        {
            _letras = _arrLetras;
            _indiceLetra = 0;
            _numeroFila = UInt32.Parse("1");
            _encabezados = new string[] {};
            _fuenteDeDatos = new DataTable();    
        }

        protected void CrearLibro()
        {
            _documentoExcel = SpreadsheetDocument.Create(_rutaArchivo, SpreadsheetDocumentType.Workbook);

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

            _workSheet = _hojaCalculo.Worksheet;
            _sheetData = _workSheet.GetFirstChild<SheetData>();   
        }

        protected virtual void AdicionarHoja(string nombreLibro, UInt32 indexTag)
        {
            Sheet hoja = new Sheet()
            {
                Id = _documentoExcel.WorkbookPart.GetIdOfPart(_hojaCalculo),
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

        public void EliminarDocumento()
        {
            if (File.Exists(_rutaArchivo))
                File.Delete(_rutaArchivo);
        }

        public void AgregarFila(OpenXmlElement hijo)
        {
            if (hijo != null)
                _sheetData.Append(hijo);
        }

        protected void CargarLetras()
        {
            int cantidadLetras = GetLonguitudColumnas();
		    int iteraciones = (cantidadLetras / 26) + 1;

            _letras = _arrLetras
                .Select((letra, indice) => new { letra, indice })
                .Where(o => o.indice < iteraciones)
                .Prepend(new { letra = "", indice = -1 })
                .Select(o => new { letra = o.letra, letras = _arrLetras })
                .SelectMany(o => o.letras, (letra1, letra2) => letra1.letra + letra2)
                .ToArray();
        }

        protected virtual void CalcularAnchoColumnas()
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match;

            var celdas = _sheetData.Descendants<Cell>();

            var configAnchoCeldas = celdas
                .GroupBy(c =>
                {
                    match = regex.Match(c.CellReference.ToString());
                    return match.Value;
                })
                .Select((g, index) => new { Letra = g.Key, Puntos = g.Max(cv => cv.CellValue.InnerText.Length) * TAMANIO_LETRA })
                .Select((o, i) => new { IndiceCelda = i + 1, o.Puntos })
                .Where(o => o.Puntos > TAMANIO_CELDA_POR_DEFECTO)
                .ToArray();

            foreach (var config in configAnchoCeldas)
            {
                EstablecerAnchoColumna(config.IndiceCelda, config.Puntos);
            }
        }

        protected void EstablecerEncabezados()
        {
            if (Encabezados == null && FuenteDeDatos != null)
            {
                var encabezados = FuenteDeDatos.Columns
                    .Cast<DataColumn>();

                if (ExcluirColumnas != null && ExcluirColumnas.Length > 0)
                    encabezados = encabezados.Where(c => !ExcluirColumnas.Contains(c.ColumnName));

                Encabezados = encabezados
                    .Select(c => c.ColumnName)
                    .ToArray();
            }
        }

        public void EstablecerAnchoColumna(int indice, DoubleValue ancho)
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

        private int GetLonguitudColumnas()
        {
            int longuitud = _encabezados.Length;

            if (longuitud == 0)
                longuitud = _fuenteDeDatos.Columns.Count;

            return longuitud;
        }

        public Row NuevaFila(int longuitud)
        {
            return NuevaFila(longuitud, CellValues.String);
        }

        public Row NuevaFila(int longuitud, EnumValue<CellValues> tipo)
        {
            return NuevaFila(longuitud, tipo, 0);
        }

        public Row NuevaFila(int longuitud, EnumValue<CellValues> tipo, UInt32Value estilo)
        {
            Row fila = GetFila();

            IEnumerable<Cell> celdas = Enumerable.Range(0, longuitud)
                .Select((indice) =>
                {
                    return new Cell
                    {
                        CellReference = _letras[indice] + fila.RowIndex,
                        DataType = tipo,
                        StyleIndex = estilo
                    };
                });

            fila.Append(celdas);
            AgregarFila(fila);

            return fila;
        }
    }
}