using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Services
{
    public class ReportFile
    {
        private static readonly string _nameDirectory = "RTM_1";

        private static readonly string _path = $@"C:\";

        private static readonly string _name = "reports.txt";

        private static readonly string _fullPath = @$"{_path}{_nameDirectory}\{_name}";

        static ReportFile()
        {
            FileDirectory.CreateDirectory(_path, _nameDirectory);
        }

        public static async Task WriteReportsAsync()
        {
            DateTime dt = DateTime.Now;
            var unix = new DateTimeOffset(dt).ToUnixTimeSeconds();

            string _message =@$"1{Environment.NewLine}{unix}";

            using StreamWriter sw = new(_fullPath, false);

            await Task.Run(() => sw.WriteLineAsync(_message));
        }
    }
}
