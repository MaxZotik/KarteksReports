using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.NetworkInformation;
using ReportsLibraryOpenXmlSdk.Services;

namespace ReportsLibraryOpenXmlSdk.Entities
{
    public struct SizePng
    {
        public long Height { get; set; }

        public long Width { get; set; }

        private readonly Logger _logger;

        public SizePng(string path) 
        {
            _logger = new Logger();
            GetSizePng(path);
        }

        private void GetSizePng(string path)
        {
            try
            {
                Bitmap png = new Bitmap(path);

                Width = long.Parse(((int)(png.Width * 9525 * 0.75)).ToString());
                Height = long.Parse(((int)(png.Height * 9525 * 0.75)).ToString());
            }
            catch (Exception ex)
            {
                _ = _logger.LogAsync(ex.Message);
            }

        }
    }
}
