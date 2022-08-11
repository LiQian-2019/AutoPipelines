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
//using EXLApplication = NetOffice.ExcelApi.Application;
//using NetOffice.ExcelApi;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

namespace AutoPipelines
{
    public class Main
    {
        //private EXLApplication excelApp = null;
        //private Workbook workbook = null;
        //private Worksheet worksheet = null;
        List<PipeLineProperty> Pipes = new List<PipeLineProperty>();

        [CommandMethod("PIPE")]
        public void Draw()
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Title = "打开属性表",
                Filter = "Excel工作簿(*.xls,*.xlsx)|*.xls;*.xlsx",
                InitialDirectory = @"D:\工程资料"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ReadPropertyTab(ofd.FileName);
                CreateEmployee();
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

        private void ReadPropertyTab(string filePathName)
        {
            Editor ed = CADApplication.DocumentManager.MdiActiveDocument.Editor;
            IWorkbook workbook = null;
            using (FileStream fs = File.Open(filePathName, FileMode.Open, FileAccess.Read))
            {
                // 2007版本
                if (filePathName.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                // 2003版本
                else if (filePathName.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                ISheet sheet;
                if (workbook.GetSheetName(0).ToLower().Contains("all"))
                {
                    sheet = workbook.GetSheetAt(0);
                    if (sheet != null && sheet.LastRowNum > 0)
                        Pipes = ReadSheetPipes(sheet);
                }
                else
                {
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        sheet = workbook.GetSheetAt(i);
                        if (sheet != null && sheet.LastRowNum > 0)
                        {
                            Pipes.AddRange(ReadSheetPipes(sheet));
                        }
                    }
                }
                ed.WriteMessage("属性表读取完成。");
            }
        }

        private List<PipeLineProperty> ReadSheetPipes(ISheet sheet)
        {
            List<PipeLineProperty> pipeEachSheet = new List<PipeLineProperty>();
            int irow = 1;
            IRow row = sheet.GetRow(irow);
            while (row != null && row.GetCell(0).CellType != NPOI.SS.UserModel.CellType.Blank)
            {
                var pipe = new PipeLineProperty
                {
                    RowInd = (ushort)(irow + 1),
                    Name = row.GetCell(0) == null ? "" : row.GetCell(0).ToString(),
                    WTName = row.GetCell(1) == null ? "" : row.GetCell(1).ToString(),
                    Connect = row.GetCell(2) == null ? "" : row.GetCell(2).ToString(),
                    Attribute = row.GetCell(3) == null ? "" : row.GetCell(3).ToString(),
                    Attachment = row.GetCell(4) == null ? "" : row.GetCell(4).ToString(),
                    X = row.GetCell(5) == null ? 0 : row.GetCell(5).NumericCellValue,
                    Y = row.GetCell(6) == null ? 0 : row.GetCell(6).NumericCellValue,
                    H = row.GetCell(7) == null ? 0 : row.GetCell(7).NumericCellValue,
                    SPH = row.GetCell(8) == null ? 0 : row.GetCell(8).NumericCellValue,
                    EPH = row.GetCell(9) == null ? 0 : row.GetCell(9).NumericCellValue,
                    WellDepth = row.GetCell(10) == null ? 0 : row.GetCell(10).NumericCellValue,
                    SPDepth = row.GetCell(11) == null ? 0 : row.GetCell(11).NumericCellValue,
                    EPDepth = row.GetCell(12) == null ? 0 : row.GetCell(12).NumericCellValue,
                    Size = row.GetCell(13) == null ? "" : row.GetCell(13).ToString(),
                    Material = row.GetCell(14) == null ? "" : row.GetCell(14).ToString(),
                    Pressure = row.GetCell(15) == null ? "" : row.GetCell(15).ToString(),
                    Voltage = row.GetCell(16) == null ? "" : row.GetCell(16).ToString(),
                    TotalBHNum = row.GetCell(17) == null ? (ushort)0 : (ushort)row.GetCell(17).NumericCellValue,
                    UsedBHNum = row.GetCell(18) == null ? (ushort)0 : (ushort)row.GetCell(18).NumericCellValue,
                    CableNum = row.GetCell(19) == null ? (ushort)0 : (ushort)row.GetCell(19).NumericCellValue,
                    Company = row.GetCell(20) == null ? "" : row.GetCell(20).ToString(),
                    BuryMethod = row.GetCell(21) == null ? "" : row.GetCell(21).ToString(),
                    BuryDate = row.GetCell(22) == null ? "" : row.GetCell(22).ToString(),
                    RoadName = row.GetCell(23) == null ? "" : row.GetCell(23).ToString(),
                    Comment = row.GetCell(24) == null ? "" : row.GetCell(24).ToString()
                };
                pipe.PipeLineType = (PipeLineType)Enum.Parse(typeof(PipeLineType), pipe.WTName.Substring(0, 2));
                pipeEachSheet.Add(pipe);
                row = sheet.GetRow(irow++);
            }
            return pipeEachSheet;
        }

        private void CreateEmployee()
        {
            Editor ed = CADApplication.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            var pipeTypes = Pipes.Select(p => p.PipeLineType).Distinct().ToList().ConvertAll(t => t.ToString());
            foreach (var type in pipeTypes)
                CreateLayer(type);

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable layers = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                foreach (var pipe in Pipes)
                {
                    // 绘制管点
                    if (string.IsNullOrEmpty(pipe.WTName)) continue;
                    if (pipe.X == 0 || pipe.Y == 0 || pipe.H == 0) continue;
                    db.Clayer = layers[pipe.PipeLineType + "P"];
                    DrawPipePoint(pipe, trans, db);

                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 绘制点名
                    db.Clayer = layers[pipe.PipeLineType + "MARK"];
                    DrawPipeName(pipe, trans, btr);

                    // 绘制管线及属性
                    if (!string.IsNullOrEmpty(pipe.Connect))
                    {
                        try
                        {
                            var endPipe = Pipes.First(p => p.WTName == pipe.Connect);
                            db.Clayer = layers[pipe.PipeLineType + "L"];
                            DrawPipeLine(pipe, endPipe, trans, btr);

                            db.Clayer = layers[pipe.PipeLineType + "T"];
                            DrawPipeText(pipe, endPipe, trans, btr);
                        }
                        catch (InvalidOperationException)
                        {

                        }

                    }

                    // 绘制高程注记
                    db.Clayer = layers[pipe.PipeLineType + "FZL"];
                    if (pipe.WellDepth > 0 || pipe.SPDepth > 0)
                    {
                        DrawPipeFZL(pipe, trans, btr);
                    }
                }
                trans.Commit();
            }
        }

        private void DrawPipeFZL(PipeLineProperty pipe, Transaction trans, BlockTableRecord btr)
        {
            Point2d startPoint = new Point2d(pipe.X, pipe.Y);
            Point2d fzlMdPoint = new Point2d(pipe.X + 3.5355, pipe.Y + 3.5355);
            Point2d fzlEdPoint = new Point2d(pipe.X + 10.5355, pipe.Y + 3.5355);
            Point3d fzlTpTextPos = new Point3d(pipe.X + 5, pipe.Y + 4.0355, pipe.H);
            Point3d fzlBtTextPos = new Point3d(pipe.X + 5, pipe.Y + 2.0355, pipe.H);
            Polyline pline = new Polyline(2);
            pline.AddVertexAt(0, startPoint, 0, 0, 0);
            pline.AddVertexAt(1, fzlMdPoint, 0, 0, 0);
            pline.AddVertexAt(2, fzlEdPoint, 0, 0, 0);

            DBText fzlTpText = new DBText
            {
                Position = fzlTpTextPos,
                TextString = pipe.H.ToString(),
                Height = 1
            };
            DBText fzlBtText = new DBText
            {
                Position = fzlBtTextPos,
                TextString = pipe.WellDepth > 0 ? (pipe.H - pipe.WellDepth).ToString() : (pipe.H - pipe.SPDepth).ToString(),
                Height = 1
            };

            btr.AppendEntity(pline);
            btr.AppendEntity(fzlTpText);
            btr.AppendEntity(fzlBtText);
            trans.AddNewlyCreatedDBObject(pline, true);
            trans.AddNewlyCreatedDBObject(fzlTpText, true);
            trans.AddNewlyCreatedDBObject(fzlBtText, true);
        }

        private void DrawPipeText(PipeLineProperty pipe, PipeLineProperty endPipe, Transaction trans, BlockTableRecord btr)
        {
            double lineLength = Math.Sqrt(Math.Pow((endPipe.X - pipe.X), 2) + Math.Pow((endPipe.Y - pipe.Y), 2) + Math.Pow((endPipe.H - pipe.H), 2));
            if (lineLength < 10) return;
            Point3d midPoint = new Point3d((pipe.X + endPipe.X) / 2, (pipe.Y + endPipe.Y) / 2, (pipe.H + endPipe.H) / 2);
            double k = (endPipe.Y - pipe.Y) / (endPipe.X - pipe.X);
            string text, kongshu;
            kongshu = pipe.TotalBHNum.ToString() + "/" + pipe.UsedBHNum.ToString();
            switch (pipe.PipeLineType)
            {
                case PipeLineType.PY:
                    text = "雨水 DN" + pipe.Size + " " + pipe.Material; break;
                case PipeLineType.WS:
                    text = "污水 DN" + pipe.Size + " " + pipe.Material; break;
                case PipeLineType.GS:
                    text = "供水 DN" + pipe.Size + " " + pipe.Material; break;
                case PipeLineType.TR:
                    text = "天然气 DN" + pipe.Size + " " + pipe.Material + " " + pipe.Pressure; break;
                case PipeLineType.GD:
                    text = "供电 " + pipe.Size + " " + pipe.Material + " " + pipe.Voltage + "kV " + kongshu + " " + pipe.CableNum.ToString() + "根"; break;
                case PipeLineType.XX:
                    text = pipe.Company + pipe.Size + " " + pipe.Material + " " + kongshu; break;
                case PipeLineType.ZM:
                    text = "路灯 " + pipe.Size + " " + pipe.Material + " " + pipe.Voltage + "kV" + " " + kongshu; break;
                case PipeLineType.ZY:
                    text = pipe.Company + " DN" + pipe.Size + " " + pipe.Material + " " + kongshu; break;
                default:
                    text = pipe.Company + " DN" + pipe.Size + " " + pipe.Material; break;
            }

            DBText pipeText = new DBText
            {
                HorizontalMode = TextHorizontalMode.TextCenter,
                Position = midPoint,
                AlignmentPoint = midPoint,
                Rotation = Math.Atan(k),
                Height = 1,
                TextString = text
            };
            btr.AppendEntity(pipeText);
            trans.AddNewlyCreatedDBObject(pipeText, true);
        }

        private void DrawPipeName(PipeLineProperty pipe, Transaction trans, BlockTableRecord btr)
        {
            DBText pipeName = new DBText
            {
                Position = new Point3d(pipe.X, pipe.Y, pipe.H),
                TextString = pipe.WTName,
                Height = 1
            };
            btr.AppendEntity(pipeName);
            trans.AddNewlyCreatedDBObject(pipeName, true);
        }

        private void DrawPipeLine(PipeLineProperty pipe, PipeLineProperty endPipe, Transaction trans, BlockTableRecord btr)
        {
            Point2d startPoint = new Point2d(pipe.X, pipe.Y);
            Point2d endPoint = new Point2d(endPipe.X, endPipe.Y);
            Polyline pline = new Polyline(1);
            pline.AddVertexAt(0, startPoint, 0, 0, 0);
            pline.AddVertexAt(1, endPoint, 0, 0, 0);
            pline.ConstantWidth = 0.1;
            btr.AppendEntity(pline);
            trans.AddNewlyCreatedDBObject(pline, true);
        }

        public void CreateLayer(string typeName)
        {
            short colorNumber;
            switch (typeName)
            {
                case "PY": colorNumber = 37; break;
                case "WS": colorNumber = 37; break;
                case "GS": colorNumber = 4; break;
                case "TR": colorNumber = 6; break;
                case "GD": colorNumber = 1; break;
                case "ZM": colorNumber = 1; break;
                case "XH": colorNumber = 1; break;
                case "XX": colorNumber = 3; break;
                case "ZY": colorNumber = 3; break;
                case "RS": colorNumber = 30; break;
                default:
                    colorNumber = 7;
                    break;
            }
            Database db = HostApplicationServices.WorkingDatabase;
            Transaction trans = db.TransactionManager.StartTransaction();
            //首先取得层表……
            LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
            //检查点号(P)层是否存在……
            if (!lt.Has(typeName + "P"))
            {
                //如果P层不存在，就创建它
                LayerTableRecord ltr = new LayerTableRecord
                {
                    Name = typeName + "P", //设置层的名字
                    Color = Color.FromColorIndex(ColorMethod.ByAci, colorNumber)
                };
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }
            //检查线(L)层是否存在……
            if (!lt.Has(typeName + "L"))
            {
                //如果L层不存在，就创建它
                LayerTableRecord ltr = new LayerTableRecord
                {
                    Name = typeName + "L", //设置层的名字
                    Color = Color.FromColorIndex(ColorMethod.ByAci, colorNumber)
                };
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }
            //检查文字(T)层是否存在……
            if (!lt.Has(typeName + "T"))
            {
                //如果T层不存在，就创建它
                LayerTableRecord ltr = new LayerTableRecord
                {
                    Name = typeName + "T", //设置层的名字
                    Color = Color.FromColorIndex(ColorMethod.ByAci, colorNumber)
                };
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }
            //检查辅助线(FZL)层是否存在……
            if (!lt.Has(typeName + "FZL"))
            {
                //如果FZL层不存在，就创建它
                LayerTableRecord ltr = new LayerTableRecord
                {
                    Name = typeName + "FZL", //设置层的名字
                    Color = Color.FromColorIndex(ColorMethod.ByAci, colorNumber)
                };
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }//检查标注(MARK)层是否存在……
            if (!lt.Has(typeName + "MARK"))
            {
                //如果MARK层不存在，就创建它
                LayerTableRecord ltr = new LayerTableRecord
                {
                    Name = typeName + "MARK", //设置层的名字
                    Color = Color.FromColorIndex(ColorMethod.ByAci, colorNumber)
                };
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }
            trans.Commit();
            trans.Dispose();
        }

        private void DrawPipePoint(PipeLineProperty pipe, Transaction trans, Database db)
        {
            string blockFileName;
            if (!string.IsNullOrEmpty(pipe.Attachment))
                blockFileName = pipe.PipeLineType + "P" + pipe.Attachment;
            else
                blockFileName = pipe.PipeLineType + "P一般管线点";
            var blockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockReference br = default;
            if (blockTable != null)
            {
                if (blockTable.Has(blockFileName))
                {
                    br = new BlockReference(new Point3d(pipe.X, pipe.Y, pipe.H), blockTable[blockFileName]);
                }
                else
                {
                    var sourceDb = new Database(false, false);
                    string blockFilePathName = "C:\\Users\\5991\\source\\repos\\LiQian-2019\\AutoPipelines\\CADBlocks\\" + blockFileName + ".dwg";
                    try
                    {
                        sourceDb.ReadDwgFile(blockFilePathName, FileShare.Read, true, "");
                    }
                    catch (System.Exception)
                    {
                        blockFileName = pipe.PipeLineType + "P一般管线点";
                        blockFilePathName = "C:\\Users\\5991\\source\\repos\\LiQian-2019\\AutoPipelines\\CADBlocks\\" + blockFileName + ".dwg";
                        sourceDb.ReadDwgFile(blockFilePathName, FileShare.Read, true, "");
                    }
                    var btrId = db.Insert(blockFileName, sourceDb, true);
                    if (!btrId.IsNull) br = new BlockReference(new Point3d(pipe.X, pipe.Y, pipe.H), btrId);
                }
            }
            var modelSpace = trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            if (modelSpace != null)
            {
                modelSpace.AppendEntity(br);
                trans.AddNewlyCreatedDBObject(br, true);
            }
        }
    }
}
