using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.App
{
    class Program
    {
        private static ArchivoExcel _excel;

        static void Main(string[] args)
        {
            DataTable dt = Datos.CrearTabla();
            _excel = new ArchivoExcel("Ejemplo.xlsx");

            try
            {
                _excel.Inicializar();
                _excel.FuenteDeDatos = dt;
                // _excel.EncabezadosHandler = ConstruirEncabezados;
                _excel.SubtitulosHandler = AgregarSubtitulos;
                /* _excel.ExcluirColumnas = new string[] { "Columna 3", "Columna 4" };
                _excel.EstilosColumnas = new EstiloColumnas[]
                {
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_NUMERICO, Columnas = new int[] { 0 } },
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_MONEDA, Columnas = new int[] { 6, 8 } },
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_FECHA, Columnas = new int[] { 7, 9 } }
                }; */
                _excel.Construir();
                _excel.Guardar();
            }
            finally
            {
                _excel.CerrarLibro();
                // excel.EliminarDocumento();
            }
        }

        static void AgregarSubtitulos()
        {
            _excel.AgregarSubTitulo("Subtitulo 1");
            _excel.AgregarSubTitulo("Subtitulo 2");
            _excel.NuevaFila(_excel.LonguitudColumnas);
        }

        static void ConstruirEncabezados()
        {
            OpenXmlElement[] openXmlElements;
            Row fila;
            string filaActual; 
            
            fila = _excel.GetFila();
            var Encabezados = new string[]
            {
                "Columnas", "",
                "Columnas", "",
                "Columnas", "",
                "Columnas", "",
                "Columnas", ""
            };

            openXmlElements = Enumerable.Range(0, Encabezados.Length)
                .Select((e, i) => new Cell {
                    CellReference = _excel.GetLetra(i) + fila.RowIndex,
                    CellValue = new CellValue(Encabezados[i]),
                    DataType = CellValues.String,
                    StyleIndex = HojaEstilos.ENCABEZADO_TABLA
                })
                .ToArray();
            
            fila.Append(openXmlElements);
            filaActual = _excel.GetNumeroFilaActual().ToString();

            _excel.AgregarFila(fila);
            _excel.CombinarCeldas("A" + filaActual, "B" + filaActual);
            _excel.CombinarCeldas("C" + filaActual, "D" + filaActual);
            _excel.CombinarCeldas("E" + filaActual, "F" + filaActual);
            _excel.CombinarCeldas("G" + filaActual, "H" + filaActual);
            _excel.CombinarCeldas("I" + filaActual, "J" + filaActual);

            fila = _excel.GetFila();
            Encabezados = new string[]
            {
                "1", "2",
                "3", "4",
                "5", "6",
                "7", "8",
                "9", "10"
            };

            openXmlElements = Enumerable.Range(0, Encabezados.Length)
                .Select((e, i) => new Cell {
                    CellReference = _excel.GetLetra(i) + fila.RowIndex,
                    CellValue = new CellValue(Encabezados[i]),
                    DataType = CellValues.String,
                    StyleIndex = HojaEstilos.ENCABEZADO_TABLA
                })
                .ToArray();
            
            fila.Append(openXmlElements);
            _excel.AgregarFila(fila);
        
        }

        static void EstablecerAnchoColumna()
        {
            _excel.SetAnchoColumna(1, 10);
        }
    }
}
