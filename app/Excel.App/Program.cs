using System;
using System.Data;
using System.Diagnostics;
using DocumentFormat.OpenXml;

namespace Excel.App
{
    class Program
    {
        private static ArchivoExcel _excel;

        static void Main(string[] args)
        {
            Stopwatch timeMeasure = new Stopwatch();
            timeMeasure.Start();

            DataTable dt = Datos.CrearTabla();
            _excel = new ArchivoExcel("Ejemplo.xlsx");

            try
            {
                _excel.Inicializar();
                _excel.FuenteDeDatos = dt;
                // _excel.EncabezadosHandler = ConstruirEncabezados;
                _excel.SubtitulosHandler = AgregarSubtitulos;
                _excel.ExcluirColumnas = new string[] { "Columna 3", "Columna 4" };
                /* _excel.EstilosColumnas = new EstiloColumnas[]
                {
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_NUMERICO, Columnas = new int[] { 0 } },
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_MONEDA, Columnas = new int[] { 8, 10 } },
                    new EstiloColumnas { Estilo = HojaEstilos.DATO_FECHA, Columnas = new int[] { 9, 11 } }
                }; */
                _excel.FilaNivelAgregado += OnFilaNivelAgregado;
                // _excel.NombreColumnaNivel = "id";
                _excel.Construir();
                _excel.Guardar();
            }
            finally
            {
                _excel.CerrarLibro();
                // excel.EliminarDocumento();
            }

            timeMeasure.Stop();
           
            TimeSpan ts = timeMeasure.Elapsed; 
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            Console.WriteLine($"Tiempo de ejecución: {elapsedTime}");
        }

        static void OnFilaNivelAgregado(object sender, FilaNivelAgregadaEventArgs e)
        {
            UInt32Value estilo =  null;
            HojaDeEstilos HojaEstilos = new HojaDeEstilos();

            if (e.Nivel == 1)
                estilo = HojaEstilos.NIVEL_UNO;
            else if (e.Nivel == 2)
                estilo = HojaEstilos.NIVEL_DOS;
            else if (e.Nivel == 3)
                estilo = HojaEstilos.NIVEL_TRES;
            else if (e.Nivel == 4)
                estilo = HojaEstilos.NIVEL_CUATRO;
            else if (e.Nivel == 5)
                estilo = HojaEstilos.NIVEL_CINCO;
            else if (e.Nivel == 6)
                estilo = HojaEstilos.NIVEL_SEIS;

            foreach(var celda in e.Celdas)
            {
                if (estilo != null)
                    celda.StyleIndex = estilo;
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
