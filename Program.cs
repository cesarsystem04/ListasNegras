using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ListaNegra
{
    class Program
    {
        static void Main(string[] args)
        {
            ProcesaArchivo();
        }


        private static void ProcesaArchivo()
        {
            string rutaArchivoEntrada = string.Empty;
            string nombreArchivoEntrada = string.Empty;
            string rutaArchivo = string.Empty;
            string msgError = string.Empty;
            string nombreArchivoSalida = string.Empty;
            List<ListaNegra> listaNegra = new List<ListaNegra>();
            bool resultadoOperacion;

            rutaArchivoEntrada = Directory.GetCurrentDirectory();
            nombreArchivoEntrada = ConfigurationManager.AppSettings["nombreArchivoEntrada"];
            rutaArchivo = string.Format("{0}\\{1}", rutaArchivoEntrada, nombreArchivoEntrada);
            nombreArchivoSalida = string.Format(@"{1}.csv", rutaArchivo, (ConfigurationManager.AppSettings["nombreArchivoSalida"] + DateTime.Now.ToString("_yyyyMMdd HHmms")));

            EscribirLog("Bitacora", "Inicio de crecion de archivo " + nombreArchivoSalida);
            resultadoOperacion = LeerArchivoEntrada(listaNegra, rutaArchivo, out msgError);

            if (resultadoOperacion)
            {
                resultadoOperacion = GeneraArchivoSalida(listaNegra, rutaArchivoEntrada, nombreArchivoSalida, out msgError);

                if (resultadoOperacion)
                {
                    EscribirLog("Bitacora", "Fin de crecion de archivo " + nombreArchivoSalida);
                }

                else
                {
                    EscribirLog("Bitacora", "Fin de crecion de archivo, error: " + msgError);
                }
            }

            else
            {
                EscribirLog("Bitacora", "Fin de crecion de archivo, error: " + msgError);
            }



        }


        private static bool LeerArchivoEntrada(List<ListaNegra> ListaNegra, string rutaArchivo, out string msgError)
        {
            msgError = string.Empty;

            try
            {

                using (FileStream fileStream = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {

                    using (SpreadsheetDocument docExcel = SpreadsheetDocument.Open(fileStream, false))
                    {

                        WorkbookPart oWorkbookPart = docExcel.WorkbookPart;
                        WorksheetPart oWorksheetPart = oWorkbookPart.WorksheetParts.First();
                        Worksheet oWorksheet = oWorksheetPart.Worksheet;

                        var sheetData = oWorksheetPart.Worksheet.Elements<SheetData>().First();

                        SharedStringTablePart sstpart = oWorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                        SharedStringTable sst = sstpart.SharedStringTable;

                        var celdas = oWorksheet.Descendants<Cell>();
                        var columnas = sheetData.Elements<Row>().ToList();

                        string strCelda;
                        string strRFC = string.Empty;
                        string strEstado = string.Empty;
                        string strFechaSAT1 = string.Empty;
                        string strFechaDOF1 = string.Empty;
                        string strFechaSAT2 = string.Empty;
                        string strFechaDOF2 = string.Empty;

                        foreach (Row columna in columnas)
                        {

                            strRFC = string.Empty;
                            strEstado = string.Empty;
                            strFechaSAT1 = string.Empty;
                            strFechaDOF1 = string.Empty;
                            strFechaSAT2 = string.Empty;
                            strFechaDOF2 = string.Empty;


                            if (columna.RowIndex > 3)
                            {

                                //foreach (Cell celda in columna.Elements<Cell>())
                                foreach (Cell celda in columna.Descendants<Cell>())
                                {
                                    strCelda = celda.CellReference.Value;

                                    string valorCelda = ObtenValorCeldaFormateado(oWorkbookPart, celda);

                                    //RFC
                                    if (celda.CellReference.Value.Equals("B" + columna.RowIndex))
                                        strRFC = valorCelda;

                                    //Situacion
                                    if (celda.CellReference.Value.Equals("D" + columna.RowIndex))
                                        //strEstado = sst.ChildElements[ssid].InnerText;
                                        strEstado = valorCelda;

                                    //Fecha SAT1
                                    if (celda.CellReference.Value.Equals("F" + columna.RowIndex))
                                        strFechaSAT1 = valorCelda;

                                    //Fecha DOT1
                                    if (celda.CellReference.Value.Equals("H" + columna.RowIndex))
                                        strFechaDOF1 = valorCelda;

                                    //Fecha SAT2
                                    if (celda.CellReference.Value.Equals("M" + columna.RowIndex))
                                        strFechaSAT2 = valorCelda;

                                    //Fecha DOT2
                                    if (celda.CellReference.Value.Equals("N" + columna.RowIndex))
                                        strFechaDOF2 = valorCelda;


                                }


                                if (strRFC != string.Empty)
                                {
                                    if (strEstado.Contains("Presunto") || strEstado.Contains("Definitivo"))
                                    {
                                        ListaNegra registroListaNegra = new ListaNegra();
                                        registroListaNegra.RFC = strRFC;
                                        registroListaNegra.Estado = strEstado;

                                        if (strEstado.Contains("Presunto"))
                                        {
                                            registroListaNegra.FechaSAT = strFechaSAT1;
                                            registroListaNegra.FechaDOF = strFechaDOF1;
                                        }
                                        else if (strEstado.Contains("Definitivo"))
                                        {
                                            registroListaNegra.FechaSAT = strFechaSAT2;
                                            registroListaNegra.FechaDOF = strFechaDOF2;
                                        }

                                        ListaNegra.Add(registroListaNegra);
                                    }
                                }

                            }

                        }
                    }

                }

            }

            catch (Exception e)
            {
                msgError = e.Message;
                return false;
            }

            return true;
        }


        private static bool GeneraArchivoSalida(List<ListaNegra> listaNegras, string rutaArchivo, string nombreArchivo, out string msgError)
        {
            bool result = false;
            msgError = string.Empty;

            try
            {

                var Cadena = new StringBuilder("");
                string separador = ",";
                string outCsvArchivo = nombreArchivo;
                var stream = File.CreateText(outCsvArchivo);

                foreach (var listn in listaNegras)
                {
                    Cadena.Append(listn.RFC + separador);
                    Cadena.Append(listn.Estado + separador);
                    Cadena.Append(listn.FechaSAT + separador);
                    Cadena.Append(listn.FechaDOF + separador);
                    Cadena.Remove(Cadena.Length - 1, 1);
                    Cadena.Append("\r\n");
                }

                stream.WriteLine(Cadena.ToString());
                stream.Close();

                result = true;

            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                return false;
            }


            return result;
        }



        private static void EscribirLog(string nombreArchivo, string mensaje)
        {
            var ruta = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            EscribirLog(ruta, nombreArchivo, mensaje);
        }

        private static void EscribirLog(string rutaArchivo, string nombreArchivo, string mensaje)
        {
            string strLogNombreArchivo = string.Empty;

            try
            {
                //Valida Directorio
                if (!System.IO.Directory.Exists(rutaArchivo))
                {
                    System.IO.Directory.CreateDirectory(rutaArchivo);
                }

                strLogNombreArchivo = rutaArchivo + "\\" + nombreArchivo + "-" + DateTime.Now.ToString("ddMMyyyy") + ".Log";
                System.IO.File.AppendAllText(strLogNombreArchivo, "[" + DateTime.Now.ToString("dd/MM/ yyyy hh:mm:ss") + "]   " + mensaje + "\r\n");


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }


        }


        /// <summary>
        /// Clase que almacena los registros de lista negra para archivo CVS
        /// </summary>
        public class ListaNegra
        {
            public string RFC { get; set; }
            public string Estado { get; set; }
            public string FechaSAT { get; set; }
            public string FechaDOF { get; set; }

        }


        /// <summary>
        /// Obtine valor en cadena de las celdas con el formato del documento
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="celda"></param>
        /// <returns>Valor de la celda</returns>
        private static string ObtenValorCeldaFormateado(WorkbookPart workbookPart, Cell celda)
        {
            string valor = string.Empty;

            if (celda == null)
            {
                return null;
            }

            else if (celda.DataType == null) //numeros y fechas
            {
                if (celda.StyleIndex != null)
                {
                    int styleIndex = (int)celda.StyleIndex.Value;
                    CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);
                    uint formatId = cellFormat.NumberFormatId.Value;

                    if (formatId == (uint)Enumeradores.Formatos.DateShort || formatId == (uint)Enumeradores.Formatos.DateLong)
                    {
                        double oaDate;
                        if (double.TryParse(celda.InnerText, out oaDate))
                        {
                            valor = DateTime.FromOADate(oaDate).ToShortDateString();
                        }
                    }
                    else
                    {
                        valor = celda.InnerText;
                    }
                }
                else
                {
                    valor = celda.InnerText;
                }


            }

            else // Strings o boolean
            {
                switch (celda.DataType.Value)
                {
                    case CellValues.SharedString:
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(celda.CellValue.InnerText));
                        valor = ssi.Text.Text;
                        break;
                    case CellValues.Boolean:
                        valor = celda.CellValue.InnerText == "0" ? "false" : "true";
                        break;
                    default:
                        valor = celda.CellValue.InnerText;
                        break;
                }
            }

            return valor;
        }

    }
}
