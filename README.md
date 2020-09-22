# CrearExcel.Csharp

Crear un archivo de Excel con la libreria DocumentFormat.OpenXml

Para mas información consultar Open XML SDK 2.5 para office disponible en [Microsoft Documentation](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk?redirectedfrom=MSDN)

### Constructores

| Tipo         | Parámetros y Tipos                                      | constructor                                                               |
|--------------|---------------------------------------------------------|---------------------------------------------------------------------------|
| ArchivoExcel | RutaArchivo: string                                     | public ArchivoExcel(string rutaArchivo)                                   |
| ArchivoExcel | RutaArchivo: string, Titulo: string                     | public ArchivoExcel(string rutaArchivo, string titulo)                    |
| ArchivoExcel | RutaArchivo: string, Titulo: string, NombreHoja: string | public ArchivoExcel(string rutaArchivo, string titulo, string nombreHoja) |

### Propiedades

| Propiedad            | Tipo                           | Descripción                                                                                            |
|----------------------|--------------------------------|--------------------------------------------------------------------------------------------------------|
| Encabezados          | string[]                       | Permite personalizar los encabezados de la tabla que se mostrarán en el arcivo de excel                |
| EstilosColumnas      | EstiloColumnas                 | Estilo para aplicar en cada celda, aplica por el indice de columna                                     |
| ExcluirColumnas      | string[]                       | Nombre de las columnas que se excluirán en la construcción del archivo de Excel                        |
| FuenteDeDatos        | DataTable                      | Objeto el cual sirve para generar los datos del arcvhivo de Excel                                      |
| HojaDeEstilos        | Stylesheet                     | Hoja de estilos para utilizar en la generación del archivo de Excel                                    |
| Letras               | string[]                       | Representa las letras de referencia, depende de la longuitud de la fuente de datos (Solo lectura)      |
| LonguitudColumnas    | int                            | Indica la cantidad de columnas que se van a generar en el objeto OpenXmlElemnt (Solo lectura)          |
| NombreColumnaNivel   | string                         | Identificador para agrupar la informacion y emite un evento para configurar el estilo que se requiera  |
| NombreHoja           | string                         | Nombre de la primera hoja de calculo del archivo de Excel                                              |
| Titulo               | string                         | Título del archivo del Excel                                                                           |
| WorkSheet            | Worksheet                      | Representa la hoja de cálculo (Solo lectura)                                                           |
| AnchoColumnasHandler | EstablecerAnchoColumnasHandler | Delegado, permite pasar una funcion como propiedad para establecer el ancho de cada columna            |
| EncabezadosHandler   | AgregarEncabezadosHandler      | Delegado, permite pasar una función como propiedad para construir encabezados complejos                |
| InformacionHandler   | AgregarInformacionHandler      | Delegado, permite pasar una función como propiedad para construir la información personalizada         |
| SubtitulosHandler    | AgregarSubtitulosHandler       | Delegado, permite pasar una función como propiedad para construir los subtítulos del archivo           |
| FilaNivelAgregado    | EventHandler                   | Evento, se decencadena cuando existe una columna que aplica como nivel de agrupación (Valor > 0)       |


### Métodos

| Nombre               | Retorno | Descripción                                                                |
|----------------------|---------|----------------------------------------------------------------------------|
| AgregarEncabezados   | void    | Agrega encabezados al documento de Excel                                   |  
| AgregarFila          | void    | Agrega una fila (OpenXmlElements) a la hoja de datos                       |
| CerrarLibro          | void    | Cierra el documento que se genero                                          |
| CombinarCeldas       | void    | Combina celdas en el archivo                                               |
| Construir            | void    | Construye el archivo de Excel                                              |
| GetFila              | Row     | Obtiene un nueva fila                                                      |
| GetLetra             | string  | Retorna la letra de acuerdo con el indice que se pasa como parámetro       |
| GetNumeroFilaActual  | UInt32  | Número de la fila referente con la direccion actual                        |
| Guardar              | void    | Guarda los elemnetos OpenXmlElements que se generaron en el WorkbookPart   |
| Inicializar          | void    | Inicializa el documento, la hoja de calculo y estilos del archivo de Excel | 
| NuevaFila            | Row     | Genera una nueva fila de acuerdo con la longuitud pasada como parámetro    |
| SetAnchoColumna      | void    | Establece el ancho de una columna                                          |     

### Modo de uso


#### Excel Básico

``` c# 

DataTable dt = new DataTable { ... };

// Instancia del objeto
var excel = new ArchivoExcel("NombreArchivo.xlsx"); 

// Inicializa el documento, la hoja de calculo y la hoja de estilos por defecto
excel.Inicializar();

// Asignación de la fuente de datos
excel.FuenteDeDatos = dt;

// Construye el archivo
excel.Construir();

// Guarda el documento generado
excel.Guardar();

// Cierra el libro
excel.CerrarLibro();

```


#### Usando otras propiedades, en una aplicación de consola

``` c#

namespace App.Console
{
    class Program
    {
        private static ArchivoExcel _excel;

        static void Main(string[] args)
        {
            DataTable dt = new DataTable { ... }
            _excel = new ArchivoExcel("Excel.xlsx");

            _excel.Inicializar();
            _excel.FuenteDeDatos = dt;
            _excel.SubtitulosHandler = AgregarSubtitulos;
            _excel.EncabezadosHandler = ConstruirEncabezados;
            _excel.ExcluirColumnas = new string[] { "Nombre Columna 1", "Nombre Columna Columna 4" };
            _excel.EstilosColumnas = new EstiloColumnas[]
            {
                new EstiloColumnas { Estilo = HojaEstilos.DATO_NUMERICO, Columnas = new int[] { 0 } },
                new EstiloColumnas { Estilo = HojaEstilos.DATO_MONEDA, Columnas = new int[] { 8, 10 } },
                new EstiloColumnas { Estilo = HojaEstilos.DATO_FECHA, Columnas = new int[] { 9, 11 } }
            };

            _excel.Construir();
            _excel.Guardar();
            _excel.CerrarLibro();
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
    }
}


```