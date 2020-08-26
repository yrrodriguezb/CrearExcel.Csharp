using System.Data;
using Excel;

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
                _excel.Construir();
                _excel.Guardar();
            }
            finally
            {
                _excel.CerrarLibro();
                // excel.EliminarDocumento();
            }
        }
    }
}
