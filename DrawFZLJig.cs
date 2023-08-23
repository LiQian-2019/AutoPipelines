using Autodesk.AutoCAD.ApplicationServices;
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
    class DrawFZLJig : DrawJig
    {
        public Polyline M_PolyLine { get; set; }
        public DBText FzlTpText { get; set; }
        public DBText FzlBtText{get;set;}
        private Point3d basePnt;
        private string topText, botText;
        private Point3d acquirePnt, endPnt;
        private Plane plane;

        public DrawFZLJig(Point3d _basePnt, string _tpText, string _btText)
        {
            basePnt = _basePnt;
            topText = _tpText;
            botText = _btText;
            plane = new Plane(Point3d.Origin, Vector3d.ZAxis);
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions optJigPoint = new JigPromptPointOptions("\n请指定高程注记的另一拐点：");
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
            endPnt = new Point3d(acquirePnt.X + 7.0710, acquirePnt.Y, acquirePnt.Z);
            Point3d txtPos = new Point3d(acquirePnt.X + 1.5, acquirePnt.Y + 0.5355, acquirePnt.Z);
            M_PolyLine = new Polyline(2);
            M_PolyLine.AddVertexAt(0, basePnt.Convert2d(plane), 0, 0, 0);
            M_PolyLine.AddVertexAt(1, acquirePnt.Convert2d(plane), 0, 0, 0);
            M_PolyLine.AddVertexAt(2, endPnt.Convert2d(plane), 0, 0, 0);
            FzlTpText = new DBText
            {
                TextString = topText,
                Height = 1,
                Position = txtPos
            };
            FzlBtText = new DBText
            {
                TextString = botText,
                Height = 1,
                Position = new Point3d(txtPos.X,txtPos.Y-2.0,txtPos.Z)
            };
            draw.Geometry.Draw(M_PolyLine);
            draw.Geometry.Draw(FzlTpText);
            draw.Geometry.Draw(FzlBtText);
            return true;
        }
    }
}
