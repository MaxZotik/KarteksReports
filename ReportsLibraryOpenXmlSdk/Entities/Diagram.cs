using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Entities
{
    abstract public class Diagram(int point, int axisXmvk, int axisXchannel, int axisYmvk, int axisYchannel, int axisZmvk, int axisZchannel)
    {
        /// <summary>
        /// Номер точки подключения
        /// </summary>
        public int Point { get; set; } = point;

        /// <summary>
        /// Номер МВК оси Х (вертикальное)
        /// </summary>
        public int AxisXmvk { get; set; } = axisXmvk;

        /// <summary>
        /// Номер канала оси Х (вертикальное)
        /// </summary>
        public int AxisXchannel { get; set; } = axisXchannel;

        /// <summary>
        /// Номер МВК оси Y (горизонтальное)
        /// </summary>
        public int AxisYmvk { get; set; } = axisYmvk;

        /// <summary>
        /// Номер канала оси Y (горизонтальное)
        /// </summary>
        public int AxisYchannel { get; set; } = axisYchannel;

        /// <summary>
        /// Номер МВК оси Z (осевое)
        /// </summary>
        public int AxisZmvk { get; set; } = axisZmvk;

        /// <summary>
        /// Номер канала оси Z (осевое)
        /// </summary>
        public int AxisZchannel { get; set; } = axisZchannel;
    }
}
