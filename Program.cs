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
                            _currValue = RemoveChars(_xlRange.Cells[i, _puntero[j] + 1].Value2.ToString());
                        }
                        else
                        {
                            _currValue = "";
                        }
                        _currRow.Add(_currValue);

                    }

                    _currRow = FormatRow(_currRow);

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
        static string DepuraProv(string _input)
        {
            string[] _cpProv = new string[] 
            {
                "Alava",
                "Albacete",
                "Alicante",
                "Almería",
                "Ávila",
                "Badajoz",
                "Baleares",
                "Barcelona",
                "Burgos",
                "Cáceres",
                "Cádiz",
                "Castellón",
                "Ciudad Real",
                "Córdoba",
                "Coruña",
                "Cuenca",
                "Gerona",
                "Granada",
                "Guadalajara",
                "Guipúzcoa",
                "Huelva",
                "Huesca",
                "Jaén",
                "León",
                "Lérida",
                "La Rioja",
                "Lugo",
                "Madrid",
                "Málaga",
                "Murcia",
                "Navarra",
                "Orense",
                "Asturias",
                "Palencia",
                "Las Palmas",
                "Pontevedra",
                "Salamanca",
                "Santa Cruz de Tenerife",
                "Cantabria",
                "Segovia",
                "Sevilla",
                "Soria",
                "Tarragona",
                "Teruel",
                "Toledo",
                "Valencia",
                "Valladolid",
                "Vizcaya",
                "Zamora",
                "Zaragoza",
                "Ceuta",
                "Melilla" };
            _input = _input.Substring(0,2);
            int _index = Convert.ToInt32(_input) - 1;
            return _cpProv[_index];
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
        static string FechaCesion(string _ignore)
        {
            DateTime _fechaCesion = DateTime.Now;
            switch (_fechaCesion.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    _fechaCesion.AddDays(2);
                    break;
                case DayOfWeek.Monday:
                    break;
                case DayOfWeek.Tuesday:
                    break;
                case DayOfWeek.Wednesday:
                    break;
                case DayOfWeek.Thursday:
                    break;
                case DayOfWeek.Friday:
                    _fechaCesion.AddDays(3);
                    break;
                case DayOfWeek.Saturday:
                    _fechaCesion.AddDays(3);
                    break;
                default:
                    break;
            }
            string _output = _fechaCesion.ToString("dd/mm/yyyy");
            return _output;
        }
        static List<string> FormatRow(List<string> _input)
        {
            List<string> _output = _input;
            //Provincia de CP
            _output[6] = DepuraProv(_output[7]);
            //Fechas
            _output[11] = DepuraFecha(_output[11]);
            _output[15] = DepuraFecha(_output[15]);
            _output[23] = DepuraFecha(_output[23]);
            //Nacionalidad
            _output[24] = Nacionalidad(_output[24]);
            //Cartera
            _output[25] = "TWINERO S.L";
            return _output;
        }
        static string Nacionalidad(string _input)
        {
            string _primeraLetra = _input.Substring(0, 1);
            switch (_primeraLetra)
            {
                case "X":
                    return "EXTRANJERO";
                case "Y":
                    return "EXTRACOMUNITARIO";
                default:
                    return "NACIONAL";
            }
        }
        static string RemoveChars (string _input)
        {
            string _output = _input.Replace(";", ",");
            _output = _output.Replace('"', '´');
            _output = _output.Replace("#", " ");
            return _output;
        }
    }
}
