using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoPipelines
{
    public class PipeLineProperty
    {
        public ushort RowInd { get; set; }
        public string Name { get; set; }
        public string WTName { get; set; }
        public PipeLineType PipeLineType { get; set; }
        public string Connect { get; set; }
        public string Attribute { get; set; }
        public string Attachment { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double H { get; set; }
        public double SPH { get; set; }
        public double EPH { get; set; }
        public double WellDepth { get; set; }
        public double SPDepth { get; set; }
        public double EPDepth { get; set; }
        public string Size { get; set; }
        public string Material { get; set; }
        public string Pressure { get; set; }
        public string Voltage { get; set; }
        public ushort TotalBHNum { get; set; }
        public ushort UsedBHNum { get; set; }
        public ushort CableNum { get; set; }
        public string Company { get; set; }
        public string BuryMethod { get; set; }
        public string BuryDate { get; set; }
        public string RoadName { get; set; }
        public string Comment { get; set; }
        public string Tag { get; set; }

    }

    public enum PipeLineType
    {
        PY,
        WS,
        GS,
        TR,
        GD,
        XX,
        ZM,
        ZY,
        XH,
        RS,
        QT
    }
}
