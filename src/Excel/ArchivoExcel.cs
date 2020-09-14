using System.Collections.Generic;
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

            var openXmlElements = encabezados
                .Select((titulo, indice) => NuevaCelda(indice, fila.RowIndex, titulo, HojaEstilos.ENCABEZADO_TABLA));
            
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
            bool trigger = columnas.Contains(NombreColumnaNivel);

            foreach(DataRow dr in filas)
            {
                fila = GetFila();
                
                xmlElemento = columnas.Cast<DataColumn>()
                    .Where(c => !ExcluirColumna(c.ColumnName))
                    .Select((c, i) => NuevaCelda(i, fila.RowIndex, dr[c.ColumnName].ToString()));
 
                fila.Append(xmlElemento);

                if (trigger && DesencadenarEventoAgregarFila(dr[NombreColumnaNivel].ToString()))
                    ConfigurarEventoFilaNivelAgregada(fila, dr[NombreColumnaNivel].ToString());
                
                AgregarFila(fila);
            }
        }

        private Cell NuevaCelda(int indice, UInt32Value fila, string texto)
        {
            return NuevaCelda(indice, fila, texto, null);
        }

        private Cell NuevaCelda(int indice, UInt32Value fila, string texto, UInt32Value estilo)
        {
            UInt32Value styleIndex = estilo;

            if (styleIndex == null)
                styleIndex = ResolverEstiloColumna(indice);

            return new Cell
            {
                CellReference = GetLetra(indice) + fila,
                DataType = ResolverTipoDeDatoCelda(texto),
                CellValue = new CellValue(texto),
                StyleIndex = styleIndex
            };
        }

        private EnumValue<CellValues> ResolverTipoDeDatoCelda(string texto)
        {
            int numeroInt = 0;
            double numeroDouble = 0;

            bool match = !texto.StartsWith("0") || texto == "0";

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

        protected bool DesencadenarEventoAgregarFila(string valorColumna)
        {
            int nivel = 0;
            int.TryParse(valorColumna, out nivel);
            return nivel > 0;
        }

        private void ConfigurarEventoFilaNivelAgregada(Row fila, string nivel)
        {
            var args = new FilaNivelAgregadaEventArgs(fila, int.Parse(nivel));
            OnFilaNivelAgregada(this, args);
        }
    }
}
