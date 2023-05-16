using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace AutoPipelines
{
    class DrawTableJig : DrawJig
    {
        private readonly Point3d basePnt;
        private readonly TypedValueList[] tVsTable;
        private Point3d acquirePnt;
        public Line TableLine { get; set; }
        public MText TextTable { get; set; }

        public DrawTableJig(Point3d _basePnt, TypedValueList[] _valueLists)
        {
            basePnt = _basePnt;
            tVsTable = _valueLists;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions optJigPoint = new JigPromptPointOptions("\n请指定表格放置的位置：");
            optJigPoint.Cursor = CursorType.Crosshair;
            optJigPoint.UserInputControls = UserInputControls.Accept3dCoordinates
                | UserInputControls.NoZeroResponseAccepted
                | UserInputControls.NoNegativeResponseAccepted;
            PromptPointResult resJigPoint = prompts.AcquirePoint(optJigPoint);
            if (resJigPoint.Status != PromptStatus.OK)
                return SamplerStatus.Cancel;
            if (resJigPoint.Value == acquirePnt)
                return SamplerStatus.NoChange;
            acquirePnt = resJigPoint.Value;
            return SamplerStatus.OK;
        }

        protected override bool WorldDraw(WorldDraw draw)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            TableLine = new Line(basePnt, acquirePnt);

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                TextTable = new MText
                {
                    Location = acquirePnt,
                    TextStyleId = AddMTextStyle("addPipeTextStyle"),
                    TextHeight = 2.0,
                    ShowBorders = true
                };

                int colWidthByte = CalColWidth();
                string contents = "管类\t\t材质\t\t" + StringTab(colWidthByte,"管径(mm)") + "埋深(m)\t压力(mPa)\t电压(kV)\t埋设方式\t权属单位\n";
                foreach (var p in tVsTable)
                {
                    contents += StringPipeType((PipeLineType)p[0].Value) + "\t\t" + p[4].Value.ToString() + "\t\t" + StringTab(colWidthByte,p[3].Value.ToString())
                        + p[1].Value.ToString() + "\t\t" + p[5].Value.ToString() + "\t\t" + p[6].Value.ToString() + "\t\t"
                        + p[7].Value.ToString() + "\t\t" + p[8].Value.ToString() + "\n";
                }
                TextTable.Contents = contents;

                draw.Geometry.Draw(TableLine);
                draw.Geometry.Draw(TextTable);
                return true;
            }
        }

        private static ObjectId AddMTextStyle(string style)
        {
            ObjectId styleId;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                DBDictionary dict = (DBDictionary)db.TableStyleDictionaryId.GetObject(OpenMode.ForRead);
                if (dict.Contains(style))
                    styleId = dict.GetAt(style);
                else
                {
                    TextStyleTableRecord ts = new TextStyleTableRecord
                    {
                        Font = new FontDescriptor("SimHei", false, false, 134, 49)
                    };
                    dict.UpgradeOpen();
                    styleId = dict.SetAt(style, ts);
                    trans.AddNewlyCreatedDBObject(ts, true);
                    trans.Commit();
                }
            }
            return styleId;
        }

        private int CalColWidth()
        {
            List<int> textSizes = new List<int>(tVsTable.Length);
            foreach (var p in tVsTable)
                textSizes.Add(Encoding.Default.GetByteCount(p[3].Value.ToString()));
            int maxSize = textSizes.Max() < 8 ? 8 : textSizes.Max();
            int colWidthByteCount = (maxSize / 6 + 1) * 6;
            return colWidthByteCount;
        }

        private string StringTab(int colWidth, string s)
        {
            int tabCount = (colWidth - Encoding.Default.GetByteCount(s) - 1) / 6 + 2;
            string sTab = s;
            for (int i = 0; i < tabCount; i++)
                sTab += "\t";
            return sTab;
        }

        private string StringPipeType(PipeLineType type)
        {
            string s;
            switch (type)
            {
                case PipeLineType.PY:
                    s = "\\C37;雨水";
                    break;
                case PipeLineType.WS:
                    s = "\\C37;污水";
                    break;
                case PipeLineType.GS:
                    s = "\\C4;给水";
                    break;
                case PipeLineType.TR:
                    s = "\\C6;燃气";
                    break;
                case PipeLineType.GD:
                    s = "\\C1;电力";
                    break;
                case PipeLineType.XX:
                    s = "\\C3;信息";
                    break;
                case PipeLineType.ZM:
                    s = "\\C1;照明";
                    break;
                case PipeLineType.ZY:
                    s = "\\C3;专用";
                    break;
                case PipeLineType.XH:
                    s = "\\C1;信号";
                    break;
                case PipeLineType.RS:
                    s = "\\C30;热力";
                    break;
                case PipeLineType.QT:
                    s = "\\C8;其它";
                    break;
                default:
                    s = "\\C7\t";
                    break;
            }
            return s;
        }
    }
}
