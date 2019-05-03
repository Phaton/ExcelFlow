using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelTwineroFlow
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadExcel(@"C:\Users\Usuario\Documents\Book1.xlsx");
            Console.ReadLine();
        }

        static public void ReadExcel(string _path)
        {
            
                Excel.Application _xlApp = new Excel.Application();
                Excel.Workbook _xlWorkbook = _xlApp.Workbooks.Open(@_path);
                Excel._Worksheet _xlWorksheet = _xlWorkbook.Sheets[1];
                Excel.Range _xlRange = _xlWorksheet.UsedRange;
                int rowCount = _xlRange.Rows.Count;
                int colCount = _xlRange.Columns.Count;
                int[] _puntero = new int[] { 0, 19, 15, 16, 22, 24, 25, 23, 21, 17, 18, 2, 3, colCount, 26, 5, 8, 9, 10, 11, colCount, colCount, 14, 20, colCount, colCount };
                int[] _punteroFun = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0 };
                Func<string, string>[] _arrayFun = new Func<string, string>[] { delegate (string s) { return LeaveAlone(s); }, delegate (string s) { return DepuraFecha(s); }, delegate (string s) { return DepuraProv(s); } };
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    List<string> _currRow = new List<string>();
                    for (int j = 0; j < _puntero.Length; j++)
                    {
                        string _currValue;
                        if (_xlRange.Cells[i, _puntero[j] + 1] != null && _xlRange.Cells[i, _puntero[j] + 1].Value2 != null)
                        {
                            _currValue = _xlRange.Cells[i, _puntero[j] + 1].Value2.ToString();
                        }
                        else
                        {
                            _currValue = "";
                        }
                        _currValue = _arrayFun[_punteroFun[j]](_currValue);
                        _currRow.Add(_currValue);

                    }
                    Console.WriteLine(String.Join(";", _currRow));
                }
            }
            finally
            {

                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(_xlRange);
                Marshal.ReleaseComObject(_xlWorksheet);

                //close and release
                _xlWorkbook.Close();
                Marshal.ReleaseComObject(_xlWorkbook);

                //quit and release
                _xlApp.Quit();
                Marshal.ReleaseComObject(_xlApp);
            }
        }

        static string DepuraFecha(string _fechain)
        {

            if (_fechain.Contains("-"))
            {
                string[] _fechasplit = _fechain.Split('-');
                return (_fechasplit[2] + "/" + _fechasplit[1] + "/" + _fechasplit[0]);

            } else if (_fechain.Contains("/"))
            {
                return _fechain;
            } else
            {
                throw new Exception("FECHA INVALIDA");
            }
        }
        static string DepuraProv(string _cpin)
        {
            string _provout = "";
            return _provout;
        }

        static string LeaveAlone(string _input)
        {
            if (_input==null)
            {
                return "";
            } else
            {
                return _input;
            }
        }

    }
}
