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
using System.Windows.Forms;
using CellType = NPOI.SS.UserModel.CellType;

namespace AutoPipelines
{
    public delegate void RowAdd(string[] rowContent);
    public delegate void BarGrow(int value);
    class AutoPipe
    {
        public List<PipeLineProperty> PipeTable { get; set; } = new List<PipeLineProperty>();
        public List<PipeLineProperty> EndPipeTable { get; set; } = new List<PipeLineProperty>();
        public string TabFilePathName { get; set; }
        private Editor Editor { get; set; }
        private Database CadDatabase { get; set; }
        private Transaction CadTransaction { get; set; }
        private LayerTable Layers { get; set; }
        internal BlockTable CadBlockTable { get; set; }
        private BlockTableRecord CadBlockTabRecord { get; set; }
        public List<string[]> ErrPipeInf { get; set; }
        private List<string> PipeTypes { get; set; }

        private bool IsDrawPipeName { get; set; }
        private bool IsDrawPipeLine { get; set; }
        private bool IsDrawPipeFzl { get; set; }
        private double PipeNameTxtHeight { get; set; }
        private int[] TextAlignMode { get; set; }
        private double ConstWidth { get; set; }
        private bool IsDisplayMaterial { get; set; }
        private bool IsDisplayPressure { get; set; }
        private bool IsDisplayVoltage { get; set; }
        private bool IsDisplaySize { get; set; }
        private bool IsDisplayBHNum { get; set; }
        private bool IsDisplayCableNum { get; set; }
        private bool IsDisplayBuryMethod { get; set; }
        private bool IsDisplayCompany { get; set; }
        private bool IsDisplayWellFzl { get; set; }
        private bool IsDisplayPointFzl { get; set; }
        private bool IsDisplayAttachmentFzl { get; set; }
        private double[] FzlAlignMode { get; set; }

        public event RowAdd AddRowValue;
        public event BarGrow AddBarValue;
        public double[] Steps { get; set; } = { 0.2, 0.4, 0.6, 0.8, 0.99, 1.0 };

        public AutoPipe()
        {
            CadDatabase = HostApplicationServices.WorkingDatabase;
        }

        public void InputConfigPara(PipeConfiguration configForm)
        {
            IsDrawPipeName = configForm.drawPipeNameChkBox.Checked;
            IsDrawPipeLine = configForm.drawPipeLineChkBox.Checked;
            IsDrawPipeFzl = configForm.drawPipeFzlChkBox.Checked;
            PipeNameTxtHeight = Convert.ToDouble(configForm.textHeightTxtBox.Text);
            switch (configForm.pipeNamePosCmbBox.SelectedIndex)
            {
                case 0: TextAlignMode = new int[2] { 2, 1 }; break;
                case 1: TextAlignMode = new int[2] { 0, 1 }; break;
                case 2: TextAlignMode = new int[2] { 2, 3 }; break;
                case 3: TextAlignMode = new int[2] { 0, 3 }; break;
                default: break;
            }
            switch (configForm.fzlPosCmbBox.SelectedIndex)
            {
                case 0: FzlAlignMode = new double[6] { -3.5355, 3.5355, -5, 4.0355, 2.0355, 2 }; break;
                case 1: FzlAlignMode = new double[6] { 3.5355, 3.5355, 5, 4.0355, 2.0355, 0 }; break;
                case 2: FzlAlignMode = new double[6] { -3.5355, -3.5355, -5, -5.0355, -3.0355, 2 }; break;
                case 3: FzlAlignMode = new double[6] { 3.5355, -3.5355, 5, -5.0355, -3.0355, 0 }; break;
                default: break;
            }

            ConstWidth = Convert.ToDouble(configForm.constWidthTxtBox.Text);
            IsDisplayMaterial = configForm.materialChkBox.Checked;
            IsDisplayPressure = configForm.pressureChkBox.Checked;
            IsDisplayVoltage = configForm.voltageChkBox.Checked;
            IsDisplaySize = configForm.sizeChkBox.Checked;
            IsDisplayBHNum = configForm.bhNumChkBox.Checked;
            IsDisplayCableNum = configForm.cableNumChkBox.Checked;
            IsDisplayBuryMethod = configForm.buryMethodChkBox.Checked;
            IsDisplayCompany = configForm.companyChkBox.Checked;
            IsDisplayWellFzl = configForm.wellFzlChkBox.Checked;
            IsDisplayPointFzl = configForm.pointFzlChkBox.Checked;
            IsDisplayAttachmentFzl = configForm.attachmentFzlChkBox.Checked;
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
                            IRow row = sheet.GetRow(0);
                            if (row.Cells[0].CellType.Equals(CellType.String) && row.Cells[0].StringCellValue.Equals("图上点号"))
                                PipeTable.AddRange(ReadSheetPipes(sheet));
                        }
                    }
                }
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
            while (row != null && row.GetCell(0).CellType != CellType.Blank)
            {
                for (int columnIndex = 0; columnIndex < 25; columnIndex++)
                {
                    if (row.GetCell(columnIndex) == null)
                        row.CreateCell(columnIndex);
                }
                try
                {
                    var pipe = new PipeLineProperty
                    {
                        RowInd = (ushort)(row.RowNum + 1),
                        Name = row.GetCell(0).CellType == CellType.String ? row.GetCell(0).StringCellValue : "",
                        WTName = row.GetCell(1).CellType == CellType.String ? row.GetCell(1).StringCellValue : "",
                        Connect = row.GetCell(2).CellType == CellType.String ? row.GetCell(2).StringCellValue : "",
                        Attribute = row.GetCell(3).CellType == CellType.String ? row.GetCell(3).StringCellValue : "",
                        Attachment = row.GetCell(4).CellType == CellType.String ? row.GetCell(4).StringCellValue : "",
                        X = row.GetCell(5).CellType == CellType.Numeric ? row.GetCell(5).NumericCellValue : 0,
                        Y = row.GetCell(6).CellType == CellType.Numeric ? row.GetCell(6).NumericCellValue : 0,
                        H = row.GetCell(7).CellType == CellType.Numeric ? row.GetCell(7).NumericCellValue : 0,
                        SPH = row.GetCell(8).CellType == CellType.Numeric ? row.GetCell(8).NumericCellValue : 0,
                        EPH = row.GetCell(9).CellType == CellType.Numeric ? row.GetCell(9).NumericCellValue : 0,
                        WellDepth = row.GetCell(10).CellType == CellType.Numeric ? row.GetCell(10).NumericCellValue : 0,
                        SPDepth = row.GetCell(11).CellType == CellType.Numeric ? row.GetCell(11).NumericCellValue : 0,
                        EPDepth = row.GetCell(12).CellType == CellType.Numeric ? row.GetCell(12).NumericCellValue : 0,
                        Size = row.GetCell(13).CellType == CellType.Blank ? "" : row.GetCell(13).ToString(),
                        Material = row.GetCell(14).CellType == CellType.Blank ? "" : row.GetCell(14).ToString(),
                        Pressure = row.GetCell(15).CellType == CellType.Blank ? "" : row.GetCell(15).ToString(),
                        Voltage = row.GetCell(16).CellType == CellType.Blank ? "" : row.GetCell(16).ToString(),
                        TotalBHNum = row.GetCell(17).CellType == CellType.Numeric ? (ushort)row.GetCell(17).NumericCellValue : (ushort)0,
                        UsedBHNum = row.GetCell(18).CellType == CellType.Numeric ? (ushort)row.GetCell(18).NumericCellValue : (ushort)0,
                        CableNum = row.GetCell(19).CellType == CellType.Numeric ? (ushort)row.GetCell(19).NumericCellValue : (ushort)0,
                        Company = row.GetCell(20).CellType == CellType.String ? row.GetCell(20).StringCellValue : "",
                        BuryMethod = row.GetCell(21).CellType == CellType.String ? row.GetCell(21).StringCellValue : "",
                        BuryDate = row.GetCell(22).CellType == CellType.String ? row.GetCell(22).StringCellValue : "",
                        RoadName = row.GetCell(23).CellType == CellType.String ? row.GetCell(23).StringCellValue : "",
                        Comment = row.GetCell(24).CellType == CellType.Blank ? "" : row.GetCell(24).ToString(),
                    };
                    pipe.PipeLineType = (PipeLineType)Enum.Parse(typeof(PipeLineType), pipe.WTName.Substring(0, 2));
                    pipeEachSheet.Add(pipe);
                    row = sheet.GetRow(++irow);
                }
                catch (Exception)
                {
                    MessageBoxButtons box = MessageBoxButtons.OK;
                    _ = MessageBox.Show($"错误发生在工作表{sheet.SheetName}，行号{irow + 1}", "提示", box);
                    row = sheet.GetRow(++irow);
                    continue;
                }
            }
            return pipeEachSheet;
        }

        /// <summary>
        /// 画管线图入口程序
        /// </summary>
        public void DrawPipes()
        {
            //adTransaction 


            // 视角转至绘图区域
            Editor = CADApplication.DocumentManager.MdiActiveDocument.Editor;
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

            using (CadTransaction = CadDatabase.TransactionManager.StartTransaction())
            {
                CadBlockTabRecord = (BlockTableRecord)CadTransaction.GetObject(CadBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                // 创建相应图层
                foreach (var type in PipeTypes)
                    CreateLayer(type);
                Layers = CadTransaction.GetObject(CadDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;

                int i = 0, j = 0;
                foreach (var pipe in PipeTable)
                {
                    if (IsDrawPipeName)
                    {
                        // 绘制管点符号
                        CadDatabase.Clayer = Layers[pipe.PipeLineType + "P"];
                        DrawPipePoint(pipe);

                        // 绘制点名
                        CadDatabase.Clayer = Layers[pipe.PipeLineType + "MARK"];
                        DrawPipeName(pipe);
                    }

                    // 绘制管线及属性
                    if (IsDrawPipeLine)
                    {
                        if (pipe.Connect.Any())
                        {
                            var endPipe = EndPipeTable[i++];
                            CadDatabase.Clayer = Layers[pipe.PipeLineType + "L"];
                            DrawPipeLine(pipe, endPipe);
                            CadDatabase.Clayer = Layers[pipe.PipeLineType + "T"];
                            DrawPipeText(pipe, endPipe);
                        }
                    }

                    // 绘制高程注记
                    if (IsDrawPipeFzl)
                    {
                        CadDatabase.Clayer = Layers[pipe.PipeLineType + "FZL"];
                        if (pipe.Attachment.Contains("井"))
                        {
                            if (IsDisplayWellFzl) DrawPipeFZL(pipe, pipe.WellDepth);
                        }
                        else if (string.IsNullOrEmpty(pipe.Attachment))
                        {
                            if (IsDisplayPointFzl) DrawPipeFZL(pipe, pipe.SPDepth);
                        }
                        else
                        {
                            if (IsDisplayAttachmentFzl) DrawPipeFZL(pipe, pipe.WellDepth);
                        }
                    }

                    if (PipeTable.IndexOf(pipe) >= PipeTable.Count() * Steps[j])
                        AddBarValue((int)(Steps[j++] * 100));
                }

                CadTransaction.Commit();
            }

            Autodesk.AutoCAD.ApplicationServices.Core.Application.UpdateScreen();
            //以下三行必须删除，否则会导致第二次执行时cad崩溃
            //CadBlockTabRecord.Dispose();
            //CadBlockTable.Dispose();
            //CadDatabase.Dispose();
        }

        /// <summary>
        /// 创建图层
        /// </summary>
        /// <param name="typeName"></param>
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
                case "QT": colorNumber = 8; break;
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

        /// <summary>
        /// 绘制管点
        /// </summary>
        /// <param name="pipe"></param>
        private void DrawPipePoint(PipeLineProperty pipe)
        {
            string blockName;
            if (pipe.Attachment.Any())
                blockName = pipe.PipeLineType + "P" + pipe.Attachment;
            else if (pipe.Attribute.Any())
                blockName = pipe.PipeLineType + "P" + pipe.Attribute;
            else
                blockName = pipe.PipeLineType + "P一般管线点";
            BlockReference br = new BlockReference(new Point3d(pipe.X, pipe.Y, pipe.H), CadBlockTable[blockName]);
            CadBlockTabRecord.AppendEntity(br);
            CadTransaction.AddNewlyCreatedDBObject(br, true);

            br.Id.AddXrecord(br.Handle.Value.ToString(), pipe.ToTypedValueList());
        }

        /// <summary>
        /// 绘制点名
        /// </summary>
        /// <param name="pipe"></param>
        private void DrawPipeName(PipeLineProperty pipe)
        {
            DBText pipeName = new DBText
            {
                HorizontalMode = (TextHorizontalMode)TextAlignMode[0],
                VerticalMode = (TextVerticalMode)TextAlignMode[1],
                AlignmentPoint = new Point3d(pipe.X, pipe.Y, pipe.H),
                Position = new Point3d(pipe.X, pipe.Y, pipe.H),
                TextString = pipe.WTName,
                Height = PipeNameTxtHeight
            };
            CadBlockTabRecord.AppendEntity(pipeName);
            CadTransaction.AddNewlyCreatedDBObject(pipeName, true);
        }

        /// <summary>
        /// 绘制连接管段
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="endPipe"></param>
        private void DrawPipeLine(PipeLineProperty pipe, PipeLineProperty endPipe)
        {
            Point2d startPoint = new Point2d(pipe.X, pipe.Y);
            Point2d endPoint = new Point2d(endPipe.X, endPipe.Y);
            Polyline pline = new Polyline(1);
            pline.AddVertexAt(0, startPoint, 0, 0, 0);
            pline.AddVertexAt(1, endPoint, 0, 0, 0);
            pline.ConstantWidth = ConstWidth;
            CadBlockTabRecord.AppendEntity(pline);
            CadTransaction.AddNewlyCreatedDBObject(pline, true);

            pline.Id.AddXrecord(pline.Handle.Value.ToString(), pipe.ToTypedValueList(method: "line"));
        }

        /// <summary>
        /// 绘制管段属性
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="endPipe"></param>
        private void DrawPipeText(PipeLineProperty pipe, PipeLineProperty endPipe)
        {
            double lineLength = Math.Sqrt(Math.Pow((endPipe.X - pipe.X), 2) + Math.Pow((endPipe.Y - pipe.Y), 2) + Math.Pow((endPipe.H - pipe.H), 2));
            if (lineLength < 10) return;
            Point3d midPoint = new Point3d((pipe.X + endPipe.X) / 2, (pipe.Y + endPipe.Y) / 2, (pipe.H + endPipe.H) / 2);
            double k = (endPipe.Y - pipe.Y) / (endPipe.X - pipe.X);
            string kongshu = pipe.TotalBHNum > 0 ? pipe.TotalBHNum.ToString() + "/" + pipe.UsedBHNum.ToString() : "";
            string dianya = pipe.Voltage == "" ? "" : pipe.Voltage + "kV";
            string guanjing = pipe.Size.Contains("BH") ? pipe.Size : pipe.Size == "" ? "" : "DN" + pipe.Size;
            string tiaoshu = pipe.CableNum > 0 ? pipe.CableNum.ToString() + "根" : "";
            string text;

            switch (pipe.PipeLineType)
            {
                case PipeLineType.PY:
                    text = "雨水" + (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "");
                    break;
                case PipeLineType.WS:
                    text = "污水" + (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "");
                    break;
                case PipeLineType.GS:
                    text = "供水" + (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "") +
                                (IsDisplayCompany ? " " + pipe.Company : "");
                    break;
                case PipeLineType.TR:
                    text = "天然气 " + (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "") +
                                (IsDisplayPressure ? " " + pipe.Pressure : "") + (IsDisplayCompany ? " " + pipe.Company : "");
                    break;
                case PipeLineType.GD:
                    text = "供电 " + (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "") +
                                (IsDisplayVoltage ? " " + dianya : "") + (IsDisplayBHNum ? " " + kongshu : "") +
                                (IsDisplayCableNum ? " " + tiaoshu : "") + (IsDisplayBuryMethod ? " " + pipe.BuryMethod : "");
                    break;
                case PipeLineType.XX:
                    text = (IsDisplayCompany ? " " + pipe.Company : "") + (IsDisplaySize ? " " + guanjing : "") +
                                (IsDisplayMaterial ? " " + pipe.Material : "") + (IsDisplayBHNum ? " " + kongshu : "") +
                                (IsDisplayCableNum ? " " + tiaoshu : "") + (IsDisplayBuryMethod ? " " + pipe.BuryMethod : "");
                    break;
                case PipeLineType.ZM:
                    text = "路灯 " + (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "") +
                                (IsDisplayVoltage ? " " + dianya : "") + (IsDisplayBHNum ? " " + kongshu : "") +
                                (IsDisplayCableNum ? " " + tiaoshu : "") + (IsDisplayBuryMethod ? " " + pipe.BuryMethod : "");
                    break;
                case PipeLineType.ZY:
                    text = (IsDisplayCompany ? " " + pipe.Company : "") + (IsDisplaySize ? " " + guanjing : "") +
                                (IsDisplayMaterial ? " " + pipe.Material : "") + (IsDisplayBHNum ? " " + kongshu : "") +
                                (IsDisplayCableNum ? " " + tiaoshu : "") + (IsDisplayBuryMethod ? " " + pipe.BuryMethod : "");
                    break;
                default:
                    text = (IsDisplaySize ? " " + guanjing : "") + (IsDisplayMaterial ? " " + pipe.Material : "") +
                                (IsDisplayBuryMethod ? " " + pipe.BuryMethod : "");
                    break;
            }

            DBText pipeText = new DBText
            {
                HorizontalMode = (TextHorizontalMode)1,
                VerticalMode = (TextVerticalMode)1,
                Position = midPoint,
                AlignmentPoint = midPoint,
                Rotation = Math.Atan(k),
                Height = 1,
                TextString = text
            };
            CadBlockTabRecord.AppendEntity(pipeText);
            CadTransaction.AddNewlyCreatedDBObject(pipeText, true);
        }

        /// <summary>
        /// 绘制标注辅助线
        /// </summary>
        /// <param name="pipe"></param>
        /// <param name="depth"></param>
        private void DrawPipeFZL(PipeLineProperty pipe, double depth)
        {
            Point2d startPoint = new Point2d(pipe.X, pipe.Y);
            Point2d fzlMdPoint = new Point2d(pipe.X + FzlAlignMode[0], pipe.Y + FzlAlignMode[1]);
            Point2d fzlEdPoint = new Point2d(pipe.X + FzlAlignMode[0] * 3, pipe.Y + FzlAlignMode[1]);
            Point3d fzlTpTextPos = new Point3d(pipe.X + FzlAlignMode[2], pipe.Y + FzlAlignMode[4], pipe.H);
            Point3d fzlBtTextPos = new Point3d(pipe.X + FzlAlignMode[2], pipe.Y + FzlAlignMode[3], pipe.H);
            Polyline pline = new Polyline(2);
            pline.AddVertexAt(0, startPoint, 0, 0, 0);
            pline.AddVertexAt(1, fzlMdPoint, 0, 0, 0);
            pline.AddVertexAt(2, fzlEdPoint, 0, 0, 0);

            DBText fzlTpText = new DBText
            {
                HorizontalMode = (TextHorizontalMode)FzlAlignMode[5],
                Position = fzlTpTextPos,
                TextString = (pipe.H - depth).ToString(),
                Height = 1
            };
            if (fzlTpText.HorizontalMode != TextHorizontalMode.TextLeft)
                fzlTpText.AlignmentPoint = fzlTpTextPos;
            DBText fzlBtText = new DBText
            {
                HorizontalMode = (TextHorizontalMode)FzlAlignMode[5],
                Position = fzlBtTextPos,
                TextString = pipe.H.ToString(),
                Height = 1
            };
            if (fzlBtText.HorizontalMode != TextHorizontalMode.TextLeft)
                fzlBtText.AlignmentPoint = fzlBtTextPos;

            CadBlockTabRecord.AppendEntity(pline);
            CadBlockTabRecord.AppendEntity(fzlTpText);
            CadBlockTabRecord.AppendEntity(fzlBtText);
            CadTransaction.AddNewlyCreatedDBObject(pline, true);
            CadTransaction.AddNewlyCreatedDBObject(fzlTpText, true);
            CadTransaction.AddNewlyCreatedDBObject(fzlBtText, true);
        }

        /// <summary>
        /// 检查属性表
        /// </summary>
        internal void CheckPropertyTab()
        {
            using (CadTransaction = CadDatabase.TransactionManager.StartTransaction())
            {
                CadBlockTable = CadTransaction.GetObject(CadDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
            }

            PipeTypes = PipeTable.Select(p => p.PipeLineType).Distinct().ToList().ConvertAll(t => t.ToString());
            ErrPipeInf = new List<string[]>();
            int i = 0;
            foreach (var pipe in PipeTable)
            {
                // 检查点名是否存在
                if (string.IsNullOrWhiteSpace(pipe.WTName))
                {
                    ErrPipeInf.Add(new string[4] { (ErrPipeInf.Count + 1).ToString(), pipe.Name, pipe.RowInd.ToString(), "缺少物探点号" });
                    AddRowValue(ErrPipeInf.Last());
                }
                // 检查坐标是否存在
                if (pipe.X == 0 || pipe.Y == 0 || pipe.H == 0)
                {
                    ErrPipeInf.Add(new string[4] { (ErrPipeInf.Count + 1).ToString(), pipe.Name, pipe.RowInd.ToString(), "缺少坐标" });
                    AddRowValue(ErrPipeInf.Last());
                }
                // 检查连接点号是否存在
                if (pipe.Connect.Any())
                {
                    try
                    {
                        var endPipe = PipeTable.First(p => p.WTName == pipe.Connect);
                        EndPipeTable.Add(endPipe);
                    }
                    catch (InvalidOperationException)
                    {
                        ErrPipeInf.Add(new string[4] { (ErrPipeInf.Count + 1).ToString(), pipe.Name, pipe.RowInd.ToString(), "未找到与之相连的点号" });
                        AddRowValue(ErrPipeInf.Last());
                    }
                }
                // 检查特征点、附属物是否存在
                string blockName;
                if (pipe.Attachment.Any())
                {
                    blockName = pipe.PipeLineType + "P" + pipe.Attachment;
                    if (!CadBlockTable.Has(blockName))
                        try
                        {
                            InsertCADBlock(blockName);
                        }
                        catch (Exception)
                        {
                            ErrPipeInf.Add(new string[4] { (ErrPipeInf.Count + 1).ToString(), pipe.Name, pipe.RowInd.ToString(), $"未找到{pipe.Attachment}对应的块文件" });
                            AddRowValue(ErrPipeInf.Last());
                        }
                }
                else if (pipe.Attribute.Any())
                {
                    blockName = pipe.PipeLineType + "P" + pipe.Attribute;
                    if (!CadBlockTable.Has(blockName))
                        try
                        {
                            InsertCADBlock(blockName);
                        }
                        catch (Exception)
                        {
                            ErrPipeInf.Add(new string[4] { (ErrPipeInf.Count + 1).ToString(), pipe.Name, pipe.RowInd.ToString(), $"未找到{pipe.Attribute}对应的块文件" });
                            AddRowValue(ErrPipeInf.Last());
                        }
                }
                else
                {
                    ErrPipeInf.Add(new string[4] { (ErrPipeInf.Count + 1).ToString(), pipe.Name, pipe.RowInd.ToString(), "缺少特征点/附属物名称" });
                    AddRowValue(ErrPipeInf.Last());
                }

                if (PipeTable.IndexOf(pipe) >= PipeTable.Count() * Steps[i])
                    AddBarValue((int)(Steps[i++] * 100));
            }
            return;
        }

        internal void InsertCADBlock(string blockName)
        {
            string blockFilePathName = Path.Combine(Environment.CurrentDirectory, "CADBlocks", blockName + ".dwg");
            var sourceDb = new Database(false, false);
            sourceDb.ReadDwgFile(blockFilePathName, FileShare.Read, true, "");
            CadDatabase.Insert(blockName, sourceDb, true);
        }
    }
}
