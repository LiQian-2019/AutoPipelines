using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.GraphicsInterface;
using NPOI.SS.Formula.Functions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
//using EXLApplication = NetOffice.ExcelApi.Application;
//using NetOffice.ExcelApi;

namespace AutoPipelines
{
    public class Main
    {
        //private EXLApplication excelApp = null;
        //private Workbook workbook = null;
        //private Worksheet worksheet = null;
        private PipeConfiguration pipeConfiguration;

        [CommandMethod("PIPE")]
        public void Draw()
        {
            if (pipeConfiguration == null)
            {
                pipeConfiguration = new PipeConfiguration();
                pipeConfiguration.ShowDialog();
            }
            else
            {
                pipeConfiguration.ShowDialog();
            }
        }

        [CommandMethod("DISPXREC")]
        public void DispXrec()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptEntityOptions opt = new PromptEntityOptions("请选择管点块");
            opt.SetRejectMessage("\n选择的不是管点块，请重新选择！");
            opt.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult entResult = ed.GetEntity(opt);
            if (entResult.Status != PromptStatus.OK) return;
            ObjectId id = entResult.ObjectId;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                ed.WriteMessage("\n" + id.GetXrecord(entResult.ObjectId.ToString()).AsArray()[15].Value.ToString());
            }
        }

        [CommandMethod("FLAG")]
        public void DrawFlag()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions opt = new PromptEntityOptions("请选择管点块");
            opt.SetRejectMessage("\n选择的不是管点块，请重新选择！");
            opt.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult entResult = ed.GetEntity(opt);
            if (entResult.Status != PromptStatus.OK) return;
            ObjectId id = entResult.ObjectId;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForWrite);
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                BlockReference br = trans.GetObject(id, OpenMode.ForRead) as BlockReference;
                LayerTable layers = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                try
                {
                    db.Clayer = layers[br.Layer.TrimEnd('P') + "FZL"];
                    // 设置高程注记文字
                    string attachment, text1, text2;
                    text2 = id.GetXrecord(id.ToString()).AsArray()[9].Value.ToString();
                    attachment = id.GetXrecord(id.ToString()).AsArray()[6].Value.ToString();
                    if (attachment != "")
                        text1 = (Convert.ToDouble(text2) - Convert.ToDouble(id.GetXrecord(id.ToString()).AsArray()[12].Value.ToString())).ToString();
                    else
                        text1 = (Convert.ToDouble(text2) - Convert.ToDouble(id.GetXrecord(id.ToString()).AsArray()[13].Value.ToString())).ToString();
                    // 实现拖拽效果
                    DrawFZLJig drawFZLJig = new DrawFZLJig(br.Position, text1, text2);
                    PromptResult pr = ed.Drag(drawFZLJig);
                    if (pr.Status == PromptStatus.OK)
                    {
                        btr.AppendEntity(drawFZLJig.M_PolyLine);
                        btr.AppendEntity(drawFZLJig.FzlTpText);
                        btr.AppendEntity(drawFZLJig.FzlBtText);
                        trans.AddNewlyCreatedDBObject(drawFZLJig.M_PolyLine, true);
                        trans.AddNewlyCreatedDBObject(drawFZLJig.FzlTpText, true);
                        trans.AddNewlyCreatedDBObject(drawFZLJig.FzlBtText, true);
                        trans.Commit();
                    }
                }
                catch { }
            }

        }


        //private void ReadPropertyTab(string fileName)
        //{
        //    Editor ed = CADApplication.DocumentManager.MdiActiveDocument.Editor;
        //    ed.WriteMessage(fileName);

        //    excelApp = new EXLApplication();
        //    workbook = excelApp.Workbooks.Open(fileName);
        //    excelApp.Visible = true;
        //    worksheet = workbook.Worksheets[1] as Worksheet;

        //    var rowsCount = worksheet.UsedRange.Rows.Count();
        //        for (ushort irow = 2; irow <= rowsCount; irow++)
        //        {
        //            var pipe = new PipeLineProperty
        //            {
        //                RowInd = irow,
        //                Name = worksheet.Cells[irow, 1].Value == null ? "" : worksheet.Cells[irow, 1].Value.ToString(),
        //                WTName = worksheet.Cells[irow, 2].Value == null ? "" : worksheet.Cells[irow, 2].Value.ToString()
        //            };
        //            pipe.PipeLineType = (PipeLineType)Enum.Parse(typeof(PipeLineType), pipe.WTName.Substring(0, 2));
        //            pipe.Connect = worksheet.Cells[irow, 3].Value == null ? "" : worksheet.Cells[irow, 3].Value.ToString();
        //            pipe.Attribute = worksheet.Cells[irow, 4].Value == null ? "" : worksheet.Cells[irow, 4].Value.ToString();
        //            pipe.Attachment = worksheet.Cells[irow, 5].Value == null ? "" : worksheet.Cells[irow, 5].Value.ToString();
        //            pipe.X = worksheet.Cells[irow, 6].Value == null ? 0 : (double)worksheet.Cells[irow, 6].Value;
        //            pipe.Y = worksheet.Cells[irow, 7].Value == null ? 0 : (double)worksheet.Cells[irow, 7].Value;
        //            pipe.H = worksheet.Cells[irow, 8].Value == null ? 0 : (double)worksheet.Cells[irow, 8].Value;
        //            pipe.SPH = worksheet.Cells[irow, 9].Value == null ? 0 : (double)worksheet.Cells[irow, 9].Value;
        //            pipe.EPH = worksheet.Cells[irow, 10].Value == null ? 0 : (double)worksheet.Cells[irow, 10].Value;
        //            pipe.WellDepth = worksheet.Cells[irow, 11].Value == null ? 0 : (double)worksheet.Cells[irow, 11].Value;
        //            pipe.SPDepth = worksheet.Cells[irow, 12].Value == null ? 0 : (double)worksheet.Cells[irow, 12].Value;
        //            pipe.EPDepth = worksheet.Cells[irow, 13].Value == null ? 0 : (double)worksheet.Cells[irow, 13].Value;
        //            pipe.Size = worksheet.Cells[irow, 14].Value == null ? "" : worksheet.Cells[irow, 14].Value.ToString();
        //            pipe.Material = worksheet.Cells[irow, 15].Value == null ? "" : worksheet.Cells[irow, 15].Value.ToString();
        //            pipe.Pressure = worksheet.Cells[irow, 16].Value == null ? "" : worksheet.Cells[irow, 16].Value.ToString();
        //            pipe.Voltage = worksheet.Cells[irow, 17].Value == null ? "" : worksheet.Cells[irow, 17].Value.ToString();
        //            pipe.TotalBHNum = (ushort)(worksheet.Cells[irow, 18].Value == null ? 0 : (ushort)worksheet.Cells[irow, 18].Value);
        //            pipe.UsedBHNum = (ushort)(worksheet.Cells[irow, 19].Value == null ? 0 : (ushort)worksheet.Cells[irow, 19].Value);
        //            pipe.CableNum = (ushort)(worksheet.Cells[irow, 20].Value == null ? 0 : (ushort)worksheet.Cells[irow, 20].Value);
        //            pipe.Company = worksheet.Cells[irow, 21].Value == null ? "" : worksheet.Cells[irow, 21].Value.ToString();
        //            pipe.BuryMethod = worksheet.Cells[irow, 22].Value == null ? "" : worksheet.Cells[irow, 22].Value.ToString();
        //            pipe.BuryDate = worksheet.Cells[irow, 23].Value == null ? "" : worksheet.Cells[irow, 23].Value.ToString();
        //            pipe.RoadName = worksheet.Cells[irow, 24].Value == null ? "" : worksheet.Cells[irow, 24].Value.ToString();
        //            pipe.Comment = worksheet.Cells[irow, 25].Value == null ? "" : worksheet.Cells[irow, 25].Value.ToString();
        //            Pipes.Add(pipe);
        //        }
        //        Console.WriteLine(pipes);
        //        ed.WriteMessage(worksheet.Cells[rowsCount-1, 1].Value.ToString());
        //        workbook.Dispose();
        //        excelApp.Visible = false;
        //        excelApp.Dispose();
        //}
    }
}
