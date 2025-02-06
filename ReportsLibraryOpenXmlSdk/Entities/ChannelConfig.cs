using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Entities
{
    public class ChannelConfig(int speedParameter, int speedFrequency, int moveParameter, int moveFrequency, int accelerateParameter, int accelerateFrequency)
    {
        public int SpeedParameter { get; set; } = speedParameter;

        public int SpeedFrequency { get; set; } = speedFrequency;

        public int MoveParameter { get; set; } = moveParameter;

        public int MoveFrequency { get; set; } = moveFrequency;

        public int AccelerateParameter { get; set; } = accelerateParameter;

        public int AccelerateFrequency { get; set; } = accelerateFrequency;
    }
}
