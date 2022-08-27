using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Hacathon.Helper
{
    public static class Helper
    {
        public static string ExcelFileName { get { return Path.Combine("Data", "HackathonTimesheet.xlsx"); } }
        public const string ExcelWorksheetName = "All Projects";


    }
}
