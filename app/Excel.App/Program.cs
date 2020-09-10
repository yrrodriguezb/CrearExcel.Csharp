using System.Data;

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
                // _excel.ExcluirColumnas = new string[] { "Columna 3", "Columna 4" };
                /* _excel.EstilosColumnas = new EstiloColumnas[]
                {
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_NUMERICO, Columnas = new int[] { 0 } },
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_MONEDA, Columnas = new int[] { 8, 10 } },
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_FECHA, Columnas = new int[] { 9, 11 } }
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
            var Encabezados = new string[]
            {
                "Columnas", "",
                "Columnas", "",
                "Columnas", "",
                "Columnas", "",
                "Columnas", ""
            };

           _excel.AgregarEncabezados(Encabezados);

            string filaActual = _excel.GetNumeroFilaActual().ToString();

            _excel.CombinarCeldas("A" + filaActual, "B" + filaActual);
            _excel.CombinarCeldas("C" + filaActual, "D" + filaActual);
            _excel.CombinarCeldas("E" + filaActual, "F" + filaActual);
            _excel.CombinarCeldas("G" + filaActual, "H" + filaActual);
            _excel.CombinarCeldas("I" + filaActual, "J" + filaActual);

            Encabezados = new string[]
            {
                "1", "2",
                "3", "4",
                "5", "6",
                "7", "8",
                "9", "10"
            };

            _excel.AgregarEncabezados(Encabezados);
        }

        static void EstablecerAnchoColumna()
        {
            _excel.SetAnchoColumna(1, 10);
        }
    }
}
