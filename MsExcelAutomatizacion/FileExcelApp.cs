using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace MsExcelAutomatizacion
{
    public class FileExcelApp: FileExcelTrace
    {
        private Excel.Application xlApp = null;
        private Excel.Workbook xlWorkBook = null;
        private Excel.Worksheet xlWorkSheet = null;
        private bool isOpen = false;
        private string _fullPathExcel = string.Empty;
        private string[] VALIDAR_EXTENSION = new string[] { "xls, xlsx" };
      
 

        public FileExcelApp(string pFullPath)
        {
            xlApp = new Excel.Application();//create a new Excel application          
            _fullPathExcel = pFullPath;
            isOpen = false;

            if (!File.Exists(_fullPathExcel))
            {
                base.setError("El archivo no existe");
                return;
            }

            string extension = Path.GetExtension(_fullPathExcel);

            if (!VALIDAR_EXTENSION.Contains(extension))
            {
                base.setError("La extensión del archivo no es válido");
                return;
            }
        }

        public void Open()
        {
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(_fullPathExcel, false, true,
                              Type.Missing, Type.Missing, Type.Missing,
                              true, Excel.XlPlatform.xlWindows, Type.Missing,
                              false, false, 0, false, true, 0);//open the workbook

                isOpen = true;
            }
            catch (Exception ex)
            {
                setError(string.Format("Error al intentar abrir el archivo {0}", _fullPathExcel), ex.Message);
                isOpen = false;
            }
        }

        public void Close()
        {
            if (!isOpen)
                return;

            xlWorkBook.Close(false, Type.Missing, Type.Missing);
            xlApp.Quit();
            //releaseObject(xlWorkSheet);
            //releaseObject(xlWorkBook);
            //releaseObject(xlApp);
            ReleaseComObjects(false);
        }

        private void ReleaseComObjects(bool isQuitting)
        {
            try
            {
                if (isQuitting)
                {
                    xlWorkBook.Close(false,Type.Missing, Type.Missing);
                    xlApp.Quit();
                }
                //Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(xlWorkBook);
                //if (worksheets != null) Marshal.ReleaseComObject(worksheets);
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlApp);
                xlWorkBook = null;
                //worksheets = null;
                xlWorkSheet = null;
                xlApp = null;
            }
            catch { }
            finally { GC.Collect(); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pColNameBegin"></param>
        /// <param name="pColNameFinish"></param>
        /// <param name="pSearchString"></param>
        /// <returns></returns>
        public int searchRowIndexInColRange(string pColNameBegin, string pColNameFinish, string pSearchString)
        {
            int rowIndex = -1;
            string strColRange = string.Format("{0}:{1}", pColNameBegin, pColNameFinish);
            Excel.Range colRange = xlWorkSheet.Columns[strColRange];//get the range object where you want to search from
            Excel.Range resultRange = colRange.Find(What: pSearchString,
                                                    LookIn: Excel.XlFindLookIn.xlValues,
                                                    LookAt: Excel.XlLookAt.xlPart,
                                                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                                                    SearchDirection: Excel.XlSearchDirection.xlNext, 
                                                    MatchCase: false,
                                                    MatchByte: Type.Missing, 
                                                    SearchFormat: Type.Missing);// search searchString in the range, if find result, return a range

            if (!(resultRange is null))
                rowIndex = resultRange.Row;

            Marshal.ReleaseComObject(colRange);
            Marshal.ReleaseComObject(resultRange);

            return rowIndex;
        }

        public string[][] getCells(string pColNameBegin, string pColNameFinish)
        {
            if (!isOpen)
            {
                base.setError("Debe inicializar la instancia a la hoja excel con el método Open");
                return null;
            }

            Excel.Range range = xlWorkSheet.get_Range(pColNameBegin, pColNameFinish);

            string[][] stringArray = null;
            object[,] cellValues;
            //string[] cells = null;
            cellValues = (range.Value2 as object[,]);

            int rowCount = cellValues.GetLength(0);
            int columnCount = cellValues.GetUpperBound(1);

            stringArray = new string[rowCount][];

            for (int index = 0; index < rowCount; index++)
            {
                stringArray[index] = new string[columnCount];

                for (int index2 = 0; index2 < columnCount; index2++)
                {
                    Object obj = cellValues.GetValue(index + 1, index2 + 1);
                    if (null != obj)
                    {
                        string value = obj.ToString();

                        stringArray[index][index2] = value;
                    }
                }
            }
           
            return stringArray;

        }

        public bool existsWorkSheetName(string WorksheetName)
        {
            if (!isOpen)
            {
                base.setError("Debe inicializar la instancia a la hoja excel con el método Open");
                return false;
            }

            
            int worksheetsCount = xlWorkBook.Worksheets.Count;
            foreach(Excel.Worksheet wsh in xlWorkBook.Worksheets)
            {
                if(string.Equals(wsh.Name, WorksheetName))
                {
                    //if(wsh.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                    return true;
                }
            }
            return false;
        }

        public void setWorksheet(string worksheetName)
        {
            if (!isOpen)
            {
                base.setError("Debe inicializar la instancia a la hoja excel con el método Open");
                return;
            }

            xlWorkSheet = xlWorkBook.Worksheets[worksheetName];//get the worksheet object
        }


    }
}
