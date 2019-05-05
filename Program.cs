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
            string folder = @"C:\Asignacion\" + DateTime.Now.ToString("yyyyMMdd") + " B1TW";
            System.IO.Directory.CreateDirectory(folder);
            string outputFile = folder + @"\" + DateTime.Now.ToString("yyyyMMdd") + "_B1TW_Cargar_"+ UltimoMiercoles()+ ".csv"; 
            ReadExcel(@"C:\Users\Usuario\Documents\Book1.xlsx", outputFile);
            Console.ReadLine();
        }

        static public void ReadExcel(string _pathInput, string _pathOutput)
        {
                //I hate Excel
                Excel.Application _xlApp = new Excel.Application();
                Excel.Workbook _xlWorkbook = _xlApp.Workbooks.Open(_pathInput);
                //Sheet index start at 1
                Excel._Worksheet _xlWorksheet = _xlWorkbook.Sheets[1];
                Excel.Range _xlRange = _xlWorksheet.UsedRange;
                int rowCount = _xlRange.Rows.Count;
                int colCount = _xlRange.Columns.Count;
                int[] _puntero = new int[] { 0, 19, 15, 16, 22, 24, 25, 23, 21, 17, 18, colCount, 2, 3, colCount, 26, 5, 8, 9, 10, 11, colCount, colCount, 14, 20, colCount, colCount };
            try
            {
                //I have to create the FIleStream so I can set the encoding to the system's default, or Excel won't read propertly special characters like "ñ"
                FileStream fs = new FileStream(_pathOutput, FileMode.Append);
                using (StreamWriter writer = new StreamWriter(fs, Encoding.Default))
                {
                    List<string> _currRow = new List<string>();
                    string _currValue = "";

                    //Add headers. Keep in mind Excel index starts on 1. First for the headers instead of an "if" statement.
                    for (int j = 0; j < _puntero.Length; j++)
                    {
                        if (_xlRange.Cells[1, _puntero[j] + 1] != null && _xlRange.Cells[1, _puntero[j] + 1].Value2 != null)
                        {
                            _currValue = RemoveChars(_xlRange.Cells[1, _puntero[j] + 1].Value2.ToString());
                        }
                        else
                        {
                            _currValue = "";
                        }
                        _currRow.Add(_currValue);
                    }
                    writer.WriteLine(String.Join(";", _currRow));


                    //Add data, row by row to _pathOutput csv file
                    for (int i = 2; i <= rowCount; i++)
                    {
                        _currRow = new List<string>();
                        for (int j = 0; j < _puntero.Length; j++)
                        {

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
                        writer.WriteLine(String.Join(";", _currRow));
                    }
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
            int fechadia;
            if (_fechain.Contains("-"))
            {
                string[] _fechasplit = _fechain.Split('-');
                return (_fechasplit[2] + "/" + _fechasplit[1] + "/" + _fechasplit[0]);

            } else if (_fechain.Contains("/"))
            {
                return _fechain;
            }
            else if (_fechain =="")
            {
                return _fechain;
            }
            //Case for the magnificient Excel's date format
            else if (Int32.TryParse(_fechain, out fechadia))
            {
                return Convert.ToDateTime("1900-01-01").AddDays(fechadia).ToString("dd/MM/yyyy");
            }
            else
            {
                throw new Exception("FECHA INVALIDA");
            }
        }
        static string DepuraProv(string _input)
        {
            //Hardcoded Postal Code-State array.
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
            _output[12] = DepuraFecha(_output[12]);
            _output[16] = DepuraFecha(_output[16]);
            _output[23] = DepuraFecha(_output[23]);
            _output[24] = DepuraFecha(_output[24]);
            //Nacionalidad
            _output[25] = Nacionalidad(_output[1]);
            //Cartera
            _output[26] = "TWINERO S.L";
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
            //I don't want those dirty characters in my db.
            string _output = _input.Replace(";", ",");
            _output = _output.Replace('"', '´');
            _output = _output.Replace("#", " ");
            return _output;
        }
        static string UltimoMiercoles()
        {
            //We should be getting Wednesday file on Friday, so if we are running this on Wednesday or Thursday, it's last week's file.
            switch (DateTime.Now.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    return DateTime.Now.AddDays(-4).ToString("yyyyMMdd");
                case DayOfWeek.Monday:
                    return DateTime.Now.AddDays(-5).ToString("yyyyMMdd");
                case DayOfWeek.Tuesday:
                    return DateTime.Now.AddDays(-6).ToString("yyyyMMdd");
                case DayOfWeek.Wednesday:
                    return DateTime.Now.AddDays(-7).ToString("yyyyMMdd");
                case DayOfWeek.Thursday:
                    return DateTime.Now.AddDays(-8).ToString("yyyyMMdd");
                case DayOfWeek.Friday:
                    return DateTime.Now.AddDays(-2).ToString("yyyyMMdd");
                case DayOfWeek.Saturday:
                    return DateTime.Now.AddDays(-3).ToString("yyyyMMdd");
                default:
                    throw new Exception("EVERYTHING BROKE");
            }
        }
    }
}
