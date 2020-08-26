using System.Data;

namespace Excel.App
{
    public class Datos
    {
        public static DataTable CrearTabla()
        {
            var columnas = 10;
            var filas = 10;

            DataTable table = new DataTable("ParentTable");
            DataColumn column;
            DataRow row;

            // Crear Columnas
            for (int i = 0; i < columnas; i++)
            {
                if (i == 0)
                {
                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.Int32");
                    column.ColumnName = "id";
                    column.ReadOnly = true;
                    column.Unique = true;
                }
                else
                {
                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = $"Columna {i}";
                    column.AutoIncrement = false;
                    column.Caption = $"Columna {i}";
                }

                table.Columns.Add(column);
                
            }

            DataColumn[] PrimaryKeyColumns = new DataColumn[1];
            PrimaryKeyColumns[0] = table.Columns["id"];
            table.PrimaryKey = PrimaryKeyColumns;

            // Crear Filas
            for (int i = 0; i <= filas; i++)
            {
                row = table.NewRow();
                row["id"] = i;
                
                for (int j = 1; j < columnas; j++)
                {
                    row[$"Columna {j}"] = $"Datos Columna [{i}][{j}]"; 
                }

                table.Rows.Add(row);
            }

            return table;
        }
    }
}