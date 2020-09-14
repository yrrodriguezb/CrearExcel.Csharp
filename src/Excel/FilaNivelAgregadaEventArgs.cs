using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public class FilaNivelAgregadaEventArgs : EventArgs
    {   
        private Row _fila { get; set; }
        public Cell[] Celdas { get; protected set; }
        public int Nivel { get; protected set; }

        public FilaNivelAgregadaEventArgs(Row fila, int nivel)
        {
            _fila = fila;
            Nivel = nivel;
            ObtenerCeldas();
        }

        private void ObtenerCeldas()
        {
            if (_fila.HasChildren) {
                Celdas = _fila.Descendants<Cell>()
                    .Select(c => {
                        c.StyleIndex = 0; 
                        return c;
                    })
                    .ToArray();
            }
        }
    }
}