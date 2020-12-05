using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public interface IHojaDeEstilos
    {
		UInt32Value TITULO { get; }
        UInt32Value SUBTITULO { get; }
        UInt32Value ENCABEZADO_TABLA { get; }
        UInt32Value DATO_NORMAL { get; }
        UInt32Value DATO_NUMERICO { get; }
        UInt32Value DATO_MONEDA { get; }
        UInt32Value DATO_FECHA { get; }
        UInt32Value NIVEL_UNO { get; }
        UInt32Value NIVEL_DOS { get; }
        UInt32Value NIVEL_TRES { get; }
        UInt32Value NIVEL_CUATRO { get; }
        UInt32Value NIVEL_CINCO { get; }
        UInt32Value NIVEL_SEIS { get; }

        /// <summary>
        /// Metodo que genera la hoja de estilos para el archivo de Excel
        /// </summary>
        /// <returns>Objeto de tipo Stylesheet.</returns>
        Stylesheet GenerarEstilos();
    }
}