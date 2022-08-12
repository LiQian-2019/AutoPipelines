using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Geometry;

namespace AutoPipelines
{
    class AutoPipe
    {
        public List<PipeLineProperty> PipeTable { get; set; } = new List<PipeLineProperty>();
        public string TabFilePathName { get; set; }
        private Editor Editor { get; set; }
        private Database CadDatabase { get; set; }
        private Transaction CadTransaction { get; set; }
        private LayerTable Layers { get; set; }
        private BlockTable CadBlockTable { get; set; }
        private BlockTableRecord CadBlockTabRecord { get; set; }
        public AutoPipe()
        {
            Editor = CADApplication.DocumentManager.MdiActiveDocument.Editor;

        }

        /// <summary>
        /// 根据给定的excel工作簿读取属性表
        /// </summary>
        public void ReadPropertyTab()
        {
            IWorkbook workbook = null;
            using (FileStream fs = File.Open(TabFilePathName, FileMode.Open, FileAccess.Read))
            {
                // 2007版本
                if (TabFilePathName.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                // 2003版本
                else if (TabFilePathName.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                ISheet sheet;
                if (workbook.GetSheetName(0).ToLower().Contains("all"))
                {
                    sheet = workbook.GetSheetAt(0);
                    if (sheet != null && sheet.LastRowNum > 0)
                        PipeTable = ReadSheetPipes(sheet);
                }
                else
                {
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        sheet = workbook.GetSheetAt(i);
                        if (sheet != null && sheet.LastRowNum > 0)
                        {
                            PipeTable.AddRange(ReadSheetPipes(sheet));
                        }
                    }
                }
                Editor.WriteMessage("属性表读取完成。");
            }
        }

        /// <summary>
        /// 根据给定的excel工作表读取属性表sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 画管线图入口程序
        /// </summary>
        public void DrawPipes()
        {
            CadDatabase = HostApplicationServices.WorkingDatabase;
            CadTransaction = CadDatabase.TransactionManager.StartTransaction();
            CadBlockTable = CadTransaction.GetObject(CadDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
            CadBlockTabRecord = (BlockTableRecord)CadTransaction.GetObject(CadBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            // 视角转至绘图区域
            ViewTableRecord viewTableRecord = Editor.GetCurrentView();
            double maxX, minX, maxY, minY;
            maxX = PipeTable.Max(p => p.X);
            minX = PipeTable.Min(p => p.X);
            maxY = PipeTable.Max(p => p.Y);
            minY = PipeTable.Min(p => p.Y);
            viewTableRecord.CenterPoint = new Point2d((maxX + minX) / 2, (maxY + minY) / 2);
            viewTableRecord.Height = maxY - minY;
            viewTableRecord.Width = maxX - minX;
            Editor.SetCurrentView(viewTableRecord);
            Autodesk.AutoCAD.ApplicationServices.Application.UpdateScreen();

            using (CadTransaction)
            {
                // 创建相应图层
                var pipeTypes = PipeTable.Select(p => p.PipeLineType).Distinct().ToList().ConvertAll(t => t.ToString());
                foreach (var type in pipeTypes)
                    CreateLayer(type);
                Layers = CadTransaction.GetObject(CadDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;

                foreach (var pipe in PipeTable)
                {
                    // 绘制管点符号
                    if (string.IsNullOrEmpty(pipe.WTName)) continue;
                    if (pipe.X == 0 || pipe.Y == 0 || pipe.H == 0) continue;
                    CadDatabase.Clayer = Layers[pipe.PipeLineType + "P"];
                    DrawPipePoint(pipe);

                    // 绘制点名
                    CadDatabase.Clayer = Layers[pipe.PipeLineType + "MARK"];
                    DrawPipeName(pipe);

                    // 绘制管线及属性
                    if (!string.IsNullOrEmpty(pipe.Connect))
                    {
                        try
                        {
                            var endPipe = PipeTable.First(p => p.WTName == pipe.Connect);
                            CadDatabase.Clayer = Layers[pipe.PipeLineType + "L"];
                            DrawPipeLine(pipe, endPipe);

                            CadDatabase.Clayer = Layers[pipe.PipeLineType + "T"];
                            DrawPipeText(pipe, endPipe);
                        }
                        catch (InvalidOperationException)
                        {

                        }

                    }

                    // 绘制高程注记
                    CadDatabase.Clayer = Layers[pipe.PipeLineType + "FZL"];
                    if (pipe.WellDepth > 0 || pipe.SPDepth > 0)
                    {
                        DrawPipeFZL(pipe);
                    }
                }

                CadTransaction.Commit();
            }
        }

        private void CreateLayer(string typeName)
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
            //首先取得层表……
            LayerTable lt = (LayerTable)CadTransaction.GetObject(CadDatabase.LayerTableId, OpenMode.ForWrite);
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
                CadTransaction.AddNewlyCreatedDBObject(ltr, true);
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
                CadTransaction.AddNewlyCreatedDBObject(ltr, true);
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
                CadTransaction.AddNewlyCreatedDBObject(ltr, true);
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
                CadTransaction.AddNewlyCreatedDBObject(ltr, true);
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
                CadTransaction.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        private void DrawPipePoint(PipeLineProperty pipe)
        {
            string blockFileName;
            if (!string.IsNullOrEmpty(pipe.Attachment))
                blockFileName = pipe.PipeLineType + "P" + pipe.Attachment;
            else
                blockFileName = pipe.PipeLineType + "P一般管线点";
            BlockReference br = default;
            if (CadBlockTable != null)
            {
                if (CadBlockTable.Has(blockFileName))
                {
                    br = new BlockReference(new Point3d(pipe.X, pipe.Y, pipe.H), CadBlockTable[blockFileName]);
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
                    var btrId = CadDatabase.Insert(blockFileName, sourceDb, true);
                    if (!btrId.IsNull) br = new BlockReference(new Point3d(pipe.X, pipe.Y, pipe.H), btrId);
                }
            }
            CadBlockTabRecord.AppendEntity(br);
            CadTransaction.AddNewlyCreatedDBObject(br, true);
        }

        private void DrawPipeName(PipeLineProperty pipe)
        {
            DBText pipeName = new DBText
            {
                Position = new Point3d(pipe.X, pipe.Y, pipe.H),
                TextString = pipe.WTName,
                Height = 1
            };
            CadBlockTabRecord.AppendEntity(pipeName);
            CadTransaction.AddNewlyCreatedDBObject(pipeName, true);
        }

        private void DrawPipeLine(PipeLineProperty pipe, PipeLineProperty endPipe)
        {
            Point2d startPoint = new Point2d(pipe.X, pipe.Y);
            Point2d endPoint = new Point2d(endPipe.X, endPipe.Y);
            Polyline pline = new Polyline(1);
            pline.AddVertexAt(0, startPoint, 0, 0, 0);
            pline.AddVertexAt(1, endPoint, 0, 0, 0);
            pline.ConstantWidth = 0.1;
            CadBlockTabRecord.AppendEntity(pline);
            CadTransaction.AddNewlyCreatedDBObject(pline, true);
        }

        private void DrawPipeText(PipeLineProperty pipe, PipeLineProperty endPipe)
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
            CadBlockTabRecord.AppendEntity(pipeText);
            CadTransaction.AddNewlyCreatedDBObject(pipeText, true);
        }

        private void DrawPipeFZL(PipeLineProperty pipe)
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

            CadBlockTabRecord.AppendEntity(pline);
            CadBlockTabRecord.AppendEntity(fzlTpText);
            CadBlockTabRecord.AppendEntity(fzlBtText);
            CadTransaction.AddNewlyCreatedDBObject(pline, true);
            CadTransaction.AddNewlyCreatedDBObject(fzlTpText, true);
            CadTransaction.AddNewlyCreatedDBObject(fzlBtText, true);
        }
    }
}
