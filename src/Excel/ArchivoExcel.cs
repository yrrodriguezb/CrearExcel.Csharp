using System.Data;
using System.Linq;
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
            Row fila = NuevaFila(LonguitudColumnas - 1, CellValues.String, 0);
            AgregarTextoFila(fila, Titulo);
        }

        public void AgregarSubTitulo(string subTitulo)
        {
            Row fila = NuevaFila(LonguitudColumnas - 1, CellValues.String, 0);
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
            Row fila = NuevaFila(LonguitudColumnas, CellValues.String, 0);

            fila.Cast<Cell>()
                .Select((celda, indice) => {
                    celda.CellValue = new CellValue(Encabezados[indice]);
                    return celda;
                })
                .ToArray();
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


            foreach(DataRow fila in filas)
            {
                filaExcel = GetFila();
                xmlElemento = new OpenXmlElement[columnas.Count];

                foreach (DataColumn columna in columnas)
                {                    
                    if (ExcluirColumna(columna.ColumnName))
                        continue;

                    celda = new Cell
                    {
                        CellReference = GetLetra(indice) + filaExcel.RowIndex,
                        DataType = CellValues.String,
                        CellValue = new CellValue(fila[columna.ColumnName].ToString()),
                        StyleIndex = 0
                    };

                    xmlElemento[indice] = celda;
                    indice++;
                }

                filaExcel.Append(xmlElemento);
                AgregarFila(filaExcel);
                indice = 0;
            }
        }
    }
}
