
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
            WorkSheet.Save();
        }

        private void AgregarTitulo()
        {
            Row fila = GetFila();
            var numeroFila = GeFilaActual().ToString();

            OpenXmlElement[] xmlElement = Enumerable.Range(-1, 15)
                .Select((indice) =>
                {
                    return new Cell
                    {
                        CellReference = GetLetra(indice + 1) + numeroFila,
                        DataType = CellValues.String,
                        CellValue = new CellValue(indice > -1 ? string.Empty : Titulo)
                    };
                })
                .ToArray();

            fila.Append(xmlElement);
            AgregarFila(fila);
            CombinarCeldas(GetLetra(0) + numeroFila, GetLetra(Letras.Length - 1).ToString() + numeroFila);
        }
    }
}
