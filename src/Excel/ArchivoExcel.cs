
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
            AgregarTitulo();

            if (SubtitulosHandler != null)
                SubtitulosHandler();

            WorkSheet.Save();
        }

        private void AgregarTextoFila(Row fila, string texto)
        {
            fila.GetFirstChild<Cell>().CellValue = new CellValue(texto);
            CombinarCeldas(GetLetra(0) + fila.RowIndex, GetLetra(LonguitudColumnas) + fila.RowIndex);
        }

        private void AgregarTitulo()
        {
            Row fila = NuevaFila(LonguitudColumnas, CellValues.String);
            AgregarTextoFila(fila, Titulo);
        }

        public void AgregarSubTitulo(string subTitulo)
        {
            Row fila = NuevaFila(LonguitudColumnas, CellValues.String, 0);
            AgregarTextoFila(fila, subTitulo);
        }
    }
}
