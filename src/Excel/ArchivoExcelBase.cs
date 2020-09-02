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

    public class EstiloColumnas
    {
        public UInt32Value Estilo { get; set; }
        public int Columna { get; set; }
        public int[] Columnas { get; set; }
    }

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
            set { _fuenteDeDatos = value; CargarLetras(); SetEncabezados(); }
        }

        private string[] _encabezados; 
        public string[] Encabezados 
        { 
            get 
            { 
                _encabezados = GetEncabezados();
                return  _encabezados; 
            } 
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
            protected set { _longuitudColumnas = value; }
        }
        
        private Worksheet _workSheet;
        public Worksheet WorkSheet
        {
            get { return _workSheet; }
            private set { _workSheet = value; }
        }

        private string[] _excluirColumnas;
        public string[] ExcluirColumnas
        {
            get { return _excluirColumnas; }
            set { _excluirColumnas = value; }
        }

        private Stylesheet _hojaDeEstilos;
        public Stylesheet HojaDeEstilos
        {
            get { return _hojaDeEstilos; }
            set { _hojaDeEstilos = SetHojaDeEstilos(value); }
        }
        
        private EstiloColumnas[] _estilosColumnas;
        public EstiloColumnas[] EstilosColumnas
        {
            get { return _estilosColumnas; }
            set { _estilosColumnas = value; AplanarConfigEstiloColumnas(); }
        }
        
        protected EstiloColumnas[] _estilos { get; set; }
        
        public string Titulo { get; set; }
        public string NombreHoja {get; set; }
        public AgregarSubtitulosHandler SubtitulosHandler { get; set; }
        public AgregarEncabezadosHandler EncabezadosHandler { get; set; }
        public AgregarInformacionHandler InformacionHandler { get; set; }
        public EstablecerAnchoColumnasHandler AnchoColumnasHandler { get; set; }


        public ArchivoExcelBase(string rutaArchivo, string titulo, string nombreHoja)
        {
            _rutaArchivo = rutaArchivo;
            Titulo = titulo;
            NombreHoja = nombreHoja;

            Validar();
            Inicializar();
        }

        private void Validar()
        {
            if (IsNull(Titulo))
                throw new NullReferenceException("El titulo no puede ser nulo");

            if (IsNull(_rutaArchivo))
                throw new NullReferenceException("La ruta para el archivo no puede ser nulo");
            
            if (IsNull(NombreHoja))
                throw new NullReferenceException("La ruta no puede ser nulo");

            if (NombreHoja.Length > 30)
                throw new InvalidOperationException("El nombre de la hoja del archivo no puede ser superior a 30 carácteres");
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

            return _letras[indice];
        }

        private void Inicializar()
        {
            _letras = _arrLetras;
            _numeroFila = UInt32.Parse("1");
            _encabezados = new string[] {};
            _excluirColumnas = new string[] {};
            _fuenteDeDatos = new DataTable();  
            _hojaDeEstilos = HojaEstilos.GenerarEstilos(); 
            _estilos = new EstiloColumnas[] {}; 
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

        protected void AdicionarHoja(string nombreLibro, UInt32 indexTag)
        {
            Sheet hoja = new Sheet()
            {
                Id = _documentoExcel.WorkbookPart.GetIdOfPart(_hojaCalculo),
                SheetId = indexTag,
                Name = nombreLibro
            };

            _hojas.Append(hoja);
        }

        protected void CrearEstilos()
        {
            WorkbookStylesPart workbookStylesPart =_documentoExcel.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = _hojaDeEstilos;
            workbookStylesPart.Stylesheet.Save();
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

            var celdas = _sheetData.Descendants<Cell>()
                .Skip(5)
                .Where(c => c.CellValue != null);

            var configAnchoColumnas = celdas
                .GroupBy(c =>
                {
                    match = regex.Match(c.CellReference.ToString());
                    return match.Value;
                })
                .Select((g, index) => new { Letra = g.Key, Puntos = g.Max(cv => cv.CellValue.InnerText.Length) * TAMANIO_LETRA })
                .Select((o, i) => new { IndiceCelda = i + 1, o.Puntos })
                .Where(o => o.Puntos > TAMANIO_CELDA_POR_DEFECTO)
                .ToArray();

            foreach (var config in configAnchoColumnas)
            {
                SetAnchoColumna(config.IndiceCelda, config.Puntos);
            }
        }

        protected void SetEncabezados()
        {
            if (Encabezados.Length == 0)
            {
                var encabezados = FuenteDeDatos.Columns
                    .Cast<DataColumn>();

                if (_excluirColumnas != null && _excluirColumnas.Length > 0)
                    encabezados = encabezados.Where(c => !_excluirColumnas.Contains(c.ColumnName));

                Encabezados = encabezados
                    .Select(c => c.ColumnName)
                    .ToArray();
            }
        }
        
        private Stylesheet SetHojaDeEstilos(Stylesheet estilos)
        {
            if (estilos != null && estilos.HasChildren)
                return estilos;

            return HojaEstilos.GenerarEstilos();
        }

        public void SetAnchoColumna(int indice, DoubleValue ancho)
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

            if (longuitud > 0 && _excluirColumnas.Length > 0)
                longuitud -= _excluirColumnas.Length;

            if (longuitud <= 0)
                longuitud = 10;

            return longuitud;
        }

        private string[] GetEncabezados()
        {
            return _encabezados
                .Where(e => !ExcluirColumna(e))
                .ToArray();
        }

        protected bool ExcluirColumna(string nombreColumna)
        {
            return ExcluirColumnas != null && ExcluirColumnas.Contains(nombreColumna);
        }

        public Row NuevaFila(int longuitud)
        {
            return NuevaFila(longuitud, "String");
        }

        public Row NuevaFila(int longuitud, string tipo)
        {
            return NuevaFila(longuitud, tipo, 0);
        }

        public Row NuevaFila(int longuitud, string tipo, UInt32Value estilo)
        {
            Row fila = GetFila();

            IEnumerable<Cell> celdas = Enumerable.Range(0, longuitud)
                .Select((indice) =>
                {
                    return new Cell
                    {
                        CellReference = _letras[indice] + fila.RowIndex,
                        DataType = GetTipoCelda(tipo),
                        StyleIndex = estilo
                    };
                });

            fila.Append(celdas);
            AgregarFila(fila);

            return fila;
        }

        protected CellValues GetTipoCelda(string nombreTipo)
        {
            Dictionary<string, CellValues> tiposCelda  = new Dictionary<string, CellValues>()
            {
                { "Date", CellValues.Date },
                { "TimeSpan", CellValues.Date },
                { "Boolean", CellValues.Boolean },
                { "Byte", CellValues.Number },
                { "Decimal", CellValues.Number },
                { "Double", CellValues.Number },
                { "Int", CellValues.Number },
                { "Char", CellValues.String },
                { "String", CellValues.String }
            };

            return tiposCelda[nombreTipo];
        }

        private void AplanarConfigEstiloColumnas()
        {
            _estilos = _estilosColumnas.SelectMany(ec => ec.Columnas, (ec, c) => new EstiloColumnas {
                Estilo = ec.Estilo,
                Columna = c
            })
            .ToArray();
        }

        private bool IsNull(string str)
        {
            return string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str);
        }
    }
}