using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLibraryOpenXmlSdk.Entities
{
    public class ArchivePoint(int point, string nameX, float axisXmovePlus, float axisXspeedPlus, float axisXacceleratePlus, float axisXmoveMinus, float axisXspeedMinus, float axisXaccelerateMinus, string nameY, float axisYmovePlus, float axisYspeedPlus, float axisYacceleratePlus, float axisYmoveMinus, float axisYspeedMinus, float axisYaccelerateMinus, string nameZ, float axisZmovePlus, float axisZspeedPlus, float axisZacceleratePlus, float axisZmoveMinus, float axisZspeedMinus, float axisZaccelerateMinus)
    {
        public int Point { get; set; } = point;


        public string NameX { get; set; } = nameX;

        public float AxisXmovePlus { get; set; } = axisXmovePlus;

        public float AxisXspeedPlus { get; set; } = axisXspeedPlus;

        public float AxisXacceleratePlus { get; set; } = axisXacceleratePlus;

        public float AxisXmoveMinus { get; set; } = axisXmoveMinus;

        public float AxisXspeedMinus { get; set; } = axisXspeedMinus;

        public float AxisXaccelerateMinus { get; set; } = axisXaccelerateMinus;



        public string NameY { get; set; } = nameY;

        public float AxisYmovePlus { get; set; } = axisYmovePlus;

        public float AxisYspeedPlus { get; set; } = axisYspeedPlus;

        public float AxisYacceleratePlus { get; set; } = axisYacceleratePlus;

        public float AxisYmoveMinus { get; set; } = axisYmoveMinus;

        public float AxisYspeedMinus { get; set; } = axisYspeedMinus;

        public float AxisYaccelerateMinus { get; set; } = axisYaccelerateMinus;




        public string NameZ { get; set; } = nameZ;

        public float AxisZmovePlus { get; set; } = axisZmovePlus;

        public float AxisZspeedPlus { get; set; } = axisZspeedPlus;

        public float AxisZacceleratePlus { get; set; } = axisZacceleratePlus;

        public float AxisZmoveMinus { get; set; } = axisZmoveMinus;

        public float AxisZspeedMinus { get; set; } = axisZspeedMinus;

        public float AxisZaccelerateMinus { get; set; } = axisZaccelerateMinus;

    }
}
