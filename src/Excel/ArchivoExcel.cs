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
            Row filaExcel = null;
            Cell celda = null;
            OpenXmlElement[] xmlElemento = null;
            int indice = 0;
            string texto = string.Empty;

            foreach(DataRow fila in filas)
            {
                filaExcel = GetFila();
                xmlElemento = new OpenXmlElement[columnas.Count];

                foreach (DataColumn columna in columnas)
                {                    
                    if (ExcluirColumna(columna.ColumnName))
                        continue;

                    texto = fila[columna.ColumnName].ToString();

                    celda = new Cell
                    {
                        CellReference = GetLetra(indice) + filaExcel.RowIndex,
                        DataType = ResolverTipoDeDatoCelda(texto),
                        CellValue = new CellValue(texto),
                        StyleIndex = ResolverEstiloColumna(indice)
                    };

                    xmlElemento[indice] = celda;
                    indice++;
                }

                filaExcel.Append(xmlElemento);
                AgregarFila(filaExcel);
                indice = 0;
                texto = string.Empty;
            }
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
