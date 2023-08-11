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
using Autodesk.AutoCAD.Colors;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;
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

        [CommandMethod("SLICETABLE")]
        public void SliceTable()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                if (!lt.Has("BMT"))
                {
                    LayerTableRecord ltr = new LayerTableRecord
                    {
                        Name = "BMT",
                        Color = Color.FromColorIndex(ColorMethod.ByColor, 0),
                    };
                    lt.Add(ltr);
                    trans.AddNewlyCreatedDBObject(ltr, true);
                }
                db.Clayer = lt["BMT"];
                trans.Commit();
            }
            PromptPointOptions optPoint = new PromptPointOptions("\n请输入第一个点：");
            PromptPointResult resPoint = ed.GetPoint(optPoint);
            if (resPoint.Status == PromptStatus.Cancel) return;
            if (resPoint.Status == PromptStatus.OK)
            {
                Point3d ptStart, ptEnd;
                ptStart = resPoint.Value;
                PromptPointOptions optPtKey = new PromptPointOptions("\n请输入下一个点：")
                {
                    UseBasePoint = true,
                    BasePoint = ptStart
                };
                PromptPointResult resKey = ed.GetPoint(optPtKey);
                ptEnd = resKey.Value;
                Line line = new Line(ptStart, ptEnd);
                db.AddToModelSpace(line);
                Point3dCollection fence = new Point3dCollection(new Point3d[] { ptStart, ptEnd });
                PromptSelectionResult result = doc.Editor.SelectFence(fence);
                List<Entity> entities = new List<Entity>();
                if (result.Status == PromptStatus.OK)
                {
                    using (Transaction trans = doc.TransactionManager.StartTransaction())
                    {
                        foreach (var id in result.Value.GetObjectIds())
                            entities.Add(trans.GetObject(id, OpenMode.ForRead) as Entity);
                        entities = entities.Take(entities.Count - 1).Filter1().ToList();
                        if (entities.Count == 0) return;
                        TypedValueList[] values = new TypedValueList[entities.Count];
                        for (int i = 0; i < values.Length; i++)
                        {
                            var xRecord = entities[i].ObjectId.GetXrecord();
                            if (xRecord == null) return;
                            values[i] = new TypedValueList(xRecord.AsArray());
                        }
                        DrawTableJig drawTableJig = new DrawTableJig(ptEnd, values);
                        PromptResult pr = ed.Drag(drawTableJig);
                        if(pr.Status == PromptStatus.OK)
                        {
                            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForWrite);
                            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                            btr.AppendEntity(drawTableJig.TableLine);
                            btr.AppendEntity(drawTableJig.TextTable);
                            trans.AddNewlyCreatedDBObject(drawTableJig.TableLine, true);
                            trans.AddNewlyCreatedDBObject(drawTableJig.TextTable, true);
                        }
                        trans.Commit();
                    }
                }
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
                    var xRecord = id.GetXrecord();
                    if (xRecord == null) return;
                    text2 = xRecord.AsArray()[9].Value.ToString();
                    attachment = xRecord.AsArray()[6].Value.ToString();
                    if (attachment != "")
                        text1 = (Convert.ToDouble(text2) - Convert.ToDouble(xRecord.AsArray()[12].Value.ToString())).ToString();
                    else
                        text1 = (Convert.ToDouble(text2) - Convert.ToDouble(xRecord.AsArray()[13].Value.ToString())).ToString();
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
                    }
                }
                catch { return; }
              trans.Commit();
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
