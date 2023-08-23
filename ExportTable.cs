using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace AutoPipelines
{
    public class ExportTable
    {
        public string SelectFlag { get; set; } = string.Empty;
        public string PipePropStr { get; set; } = string.Empty;
        public string FilePathName { get; set; } = string.Empty;

        private Editor Editor { get; set; }
        private Database CadDatabase { get; set; }
        public Document CadDocument { get; set; }

        private TypedValueList values { get; set; }
        public List<PipeLineProperty> PipeTable { get; set; }

        public ExportTable()
        {
            CadDocument = CADApplication.DocumentManager.MdiActiveDocument;
            CadDatabase = HostApplicationServices.WorkingDatabase;
            Editor = CadDocument.Editor;

            values = new TypedValueList
            {
                new TypedValue((int)DxfCode.Operator,"<and"),
                new TypedValue((int)DxfCode.Start,"Insert")
            };
        }

        internal bool ExcuteExport()
        {
            PromptSelectionResult psr;
            if (SelectFlag.Length * PipePropStr.Length == 0)
                return false;
            switch (PipePropStr)
            {
                case "ALL":
                    break;
                default:
                    string[] pipeTypes = PipePropStr.Split(',');
                    values.Add(DxfCode.Operator, "<or");
                    foreach (var type in pipeTypes)
                    {
                        values.Add(new TypedValue((int)DxfCode.LayerName, type + "P"));
                    }
                    values.Add(DxfCode.Operator, "or>");
                    break;
            }
            values.Add(DxfCode.Operator, "and>");
            switch (SelectFlag)
            {
                case "ALL":
                    psr = Editor.SelectAll(new SelectionFilter(values));
                    break;
                case "USERSELECT":
                    psr = Editor.GetSelection(new SelectionFilter(values));
                    break;
                default:
                    return false;
            }
            if (psr.Status == PromptStatus.None)
            {
                CADApplication.ShowAlertDialog("未选中任何管点。");
                return false;
            }
            SelectionSet ss = psr.Value;

            using (Transaction trans = CadDatabase.TransactionManager.StartTransaction())
            {
                PipeTable = new List<PipeLineProperty>();
                TypedValueList  typedValues;
                foreach (var pId in ss.GetObjectIds())
                {
                    if (pId.GetXrecord() == null) continue;
                    typedValues = pId.GetXrecord();
                    PipeTable.Add(typedValues.ToPipeLineProperty());
                }
            }
            
            WritePropertyTab(FilePathName);
            return true;
            //CADApplication.ShowAlertDialog("属性表导出完成。");
        }

        private void WritePropertyTab(string filePathName)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("ALL");
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            IFont font = workbook.CreateFont();
            font.IsBold = true;
            style.SetFont(font);

            // 表头
            IRow row0 = sheet.CreateRow(0);
            for (int i = 0; i < 25; i++)
            {
                row0.CreateCell(i);
                row0.GetCell(i).CellStyle = style;
            }
            row0.Cells[0].SetCellValue("图上点号");
            row0.Cells[1].SetCellValue("物探点号");
            row0.Cells[2].SetCellValue("连接点号");
            row0.Cells[3].SetCellValue("特征点");
            row0.Cells[4].SetCellValue("附属物名称");
            row0.Cells[5].SetCellValue("X坐标");
            row0.Cells[6].SetCellValue("Y坐标");
            row0.Cells[7].SetCellValue("地面高程");
            row0.Cells[8].SetCellValue("起始管线点高程");
            row0.Cells[9].SetCellValue("终止管线点高程");
            row0.Cells[10].SetCellValue("井深");
            row0.Cells[11].SetCellValue("起始管线点埋深");
            row0.Cells[12].SetCellValue("终止管线点埋深");
            row0.Cells[13].SetCellValue("管径或断面尺寸");
            row0.Cells[14].SetCellValue("材质");
            row0.Cells[15].SetCellValue("压力");
            row0.Cells[16].SetCellValue("电压");
            row0.Cells[17].SetCellValue("总孔数");
            row0.Cells[18].SetCellValue("已用孔数");
            row0.Cells[19].SetCellValue("电缆条数");
            row0.Cells[20].SetCellValue("权属单位");
            row0.Cells[21].SetCellValue("埋设方式");
            row0.Cells[22].SetCellValue("埋设日期");
            row0.Cells[23].SetCellValue("道路名称");
            row0.Cells[24].SetCellValue("备注");

            // 表格内容
            font.IsBold = false;
            style.SetFont(font);
            for (int i = 0; i < PipeTable.Count(); i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < 25; j++)
                {
                    row.CreateCell(j);
                    row.GetCell(j).CellStyle = style;
                }
                row.Cells[0].SetCellValue(PipeTable[i].Name);
                row.Cells[1].SetCellValue(PipeTable[i].WTName);
                row.Cells[2].SetCellValue(PipeTable[i].Connect);
                row.Cells[3].SetCellValue(PipeTable[i].Attribute);
                row.Cells[4].SetCellValue(PipeTable[i].Attachment);
                row.Cells[5].SetCellValue(PipeTable[i].X);
                row.Cells[6].SetCellValue(PipeTable[i].Y);
                row.Cells[7].SetCellValue(PipeTable[i].H);
                row.Cells[8].SetCellValue(PipeTable[i].SPH);
                row.Cells[9].SetCellValue(PipeTable[i].EPH);
                row.Cells[10].SetCellValue(PipeTable[i].WellDepth);
                row.Cells[11].SetCellValue(PipeTable[i].SPDepth);
                row.Cells[12].SetCellValue(PipeTable[i].EPDepth);
                row.Cells[13].SetCellValue(PipeTable[i].Size);
                row.Cells[14].SetCellValue(PipeTable[i].Material);
                row.Cells[15].SetCellValue(PipeTable[i].Pressure);
                row.Cells[16].SetCellValue(PipeTable[i].Voltage);
                row.Cells[17].SetCellValue(PipeTable[i].TotalBHNum);
                row.Cells[18].SetCellValue(PipeTable[i].UsedBHNum);
                row.Cells[19].SetCellValue(PipeTable[i].CableNum);
                row.Cells[20].SetCellValue(PipeTable[i].Company);
                row.Cells[21].SetCellValue(PipeTable[i].BuryMethod);
                row.Cells[22].SetCellValue(PipeTable[i].BuryDate);
                row.Cells[23].SetCellValue(PipeTable[i].RoadName);
                row.Cells[24].SetCellValue(PipeTable[i].Comment);
            }

            using (FileStream fs = new FileStream(filePathName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        private void CheckInputPipeProp(string inputStr)
        {
            var pipeProps = inputStr.Split(',');
            foreach (var p in pipeProps)
            {
                Enum.IsDefined(typeof(PipeLineType), p);
            }
        }
    }
}
