using ReportsLibraryOpenXmlSdk.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Services
{
    public class Logger : ILogger
    {
        private static readonly string _nameDirectory = "Log";

        private static readonly string _path = AppDomain.CurrentDomain.BaseDirectory;

        private static readonly string _name = "_log.txt";

        private static readonly string _fullPath = @$"{_path}{_nameDirectory}\{_name}";

        static Logger()
        {
            FileDirectory.CreateDirectory(_path, _nameDirectory);
        }

        /// <summary>
        /// Метод для записи логов в файл
        /// </summary>
        /// <param name="message">Сообщение</param>
        public async Task LogAsync(string message)
        {
            string _message = $@"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} : {message}";

            await WriteLogAsync(_message);
        }

        private static async Task WriteLogAsync(string message)
        {
            using StreamWriter sw = new(_fullPath, true, System.Text.Encoding.UTF8);

            await Task.Run(() => sw.WriteLineAsync(message));
        }
    }
}
