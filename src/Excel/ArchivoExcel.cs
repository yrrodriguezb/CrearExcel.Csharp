using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public class ArchivoExcel : ArchivoExcelBase
    {
        public ArchivoExcel(string rutaArchivo) : this(rutaArchivo, "Titulo Excel") { }
        public ArchivoExcel(string rutaArchivo, string titulo) : this(rutaArchivo, titulo, "Informe") { }
        public ArchivoExcel(string rutaArchivo, string titulo, string nombreHoja) : base(rutaArchivo, titulo, nombreHoja) { }

        public void Inicializar()
        {
            CrearLibro();
            CrearEstilos();
            CrearHoja(NombreHoja, 1);
        }

        public void Construir()
        {
            SetHandlers();
            AgregarTitulo();
            SubtitulosHandler();
            EncabezadosHandler();
            InformacionHandler();
            AnchoColumnasHandler();
            WorkSheet.Save();
        }

        private void AgregarTextoFila(Row fila, string texto)
        {
            fila.GetFirstChild<Cell>().CellValue = new CellValue(texto);
            CombinarCeldas(GetLetra(0) + fila.RowIndex, GetLetra(LonguitudColumnas - 1) + fila.RowIndex);
        }

        private void AgregarTitulo()
        {
            Row fila = NuevaFila(LonguitudColumnas - 1, "String", HojaEstilos.TITULO);
            AgregarTextoFila(fila, Titulo);
        }

        public void AgregarSubTitulo(string subTitulo)
        {
            Row fila = NuevaFila(LonguitudColumnas - 1, "String", HojaEstilos.SUBTITULO);
            AgregarTextoFila(fila, subTitulo);
        }

        private void SetHandlers()
        {
            if (EncabezadosHandler == null)
                EncabezadosHandler = AgregarEncabezadosTabla;

            if (SubtitulosHandler == null)
                SubtitulosHandler = AgregarFilaVacia;

            if (InformacionHandler == null)
                InformacionHandler = AgregarInformacion;

            if (AnchoColumnasHandler == null)
                AnchoColumnasHandler = CalcularAnchoColumnas;
        }

        private void AgregarEncabezadosTabla()
        {
            AgregarEncabezados(Encabezados);
        }

        public void AgregarEncabezados(string[] encabezados)
        {
            if (encabezados == null || encabezados.Length == 0)
                return;

            Row fila = GetFila();

            OpenXmlElement[] openXmlElements = encabezados
                .Select((titulo, indice) => new Cell {
                    CellReference = GetLetra(indice) + fila.RowIndex,
                    CellValue = new CellValue(titulo),
                    DataType = CellValues.String,
                    StyleIndex = HojaEstilos.ENCABEZADO_TABLA
                })
                .ToArray();
            
            fila.Append(openXmlElements);
            AgregarFila(fila);
        }

        private void AgregarFilaVacia()
        {
             Row fila = NuevaFila(LonguitudColumnas);
             AgregarTextoFila(fila, string.Empty);
        }

        private void AgregarInformacion()
        {
            DataRowCollection filas = FuenteDeDatos.Rows;
            DataColumnCollection columnas = FuenteDeDatos.Columns;
            Row fila = null;
            IEnumerable<Cell> xmlElemento = null;

            foreach(DataRow dr in filas)
            {
                fila = GetFila();
            
                xmlElemento = columnas
                    .Cast<DataColumn>()
                    .Where(c => !ExcluirColumna(c.ColumnName))
                    .Select((c, i) => NuevaCelda(i, fila.RowIndex, dr[c.ColumnName].ToString()));

                fila.Append(xmlElemento);
                AgregarFila(fila);
            }
        }

        private Cell NuevaCelda(int indice, UInt32Value fila, string texto)
        {
            return new Cell
            {
                CellReference = GetLetra(indice) + fila,
                DataType = ResolverTipoDeDatoCelda(texto),
                CellValue = new CellValue(texto),
                StyleIndex = ResolverEstiloColumna(indice)
            };
        }

        private EnumValue<CellValues> ResolverTipoDeDatoCelda(string texto)
        {
            int numeroInt = 0;
            double numeroDouble = 0;

            bool match = new Regex(@"^0$|^[^0+]\d+(.\d+)?$").Match(texto).Success;

            if ((int.TryParse(texto, out numeroInt) || double.TryParse(texto, out numeroDouble)) && match)
                return GetTipoCelda("Int");

            return GetTipoCelda("String");
        }

        private UInt32Value ResolverEstiloColumna(int indiceColumna)
        {
            UInt32Value estilo = HojaEstilos.DATO_NORMAL;

            if (_estilos.Count() > 0)
            {
                var estiloColumna = _estilos.Where(ec => ec.Columna == indiceColumna).FirstOrDefault();

                if (estiloColumna != null)
                    estilo = estiloColumna.Estilo;
            }

            return estilo;
        }
    }
}
