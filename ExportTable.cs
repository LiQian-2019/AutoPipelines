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

        public SelectionSet SelectedPoints { get; set; }
        private TypedValueList FilterInsert { get; set; }
        public List<PipeLineProperty> PipeTable { get; set; }

        public ExportTable()
        {
            CadDocument = CADApplication.DocumentManager.MdiActiveDocument;
            CadDatabase = HostApplicationServices.WorkingDatabase;
            Editor = CadDocument.Editor;

            FilterInsert = new TypedValueList
            {
                new TypedValue((int)DxfCode.Start,"Insert")
            };
        }

        /// <summary>
        /// 执行导出到Excel的方法
        /// </summary>
        /// <param name="ss"></param>
        /// <returns></returns>
        internal bool ExcuteExport(SelectionSet ss)
        {
            using (Transaction trans = CadDatabase.TransactionManager.StartTransaction())
            {
                List<Entity> entities = new List<Entity>();
                foreach (var id in ss.GetObjectIds())
                    entities.Add(trans.GetObject(id, OpenMode.ForRead) as Entity);
                if(PipePropStr != "ALLTYPES")
                    entities = entities.Filter2(PipePropStr.Split(',')).Reverse().ToList();
                if (entities.Count == 0) return false;

                PipeTable = new List<PipeLineProperty>();
                TypedValueList  typedValues;
                foreach (var entity in entities)
                {
                    if (entity.Id.GetXrecord() == null) continue;
                    typedValues = entity.Id.GetXrecord();
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
                // 此处需要倒着读取PipeTable，Excel中才能正顺序写
                int k = PipeTable.Count() - i - 1;
                row.Cells[0].SetCellValue(PipeTable[k].Name);
                row.Cells[1].SetCellValue(PipeTable[k].WTName);
                row.Cells[2].SetCellValue(PipeTable[k].Connect);
                row.Cells[3].SetCellValue(PipeTable[k].Attribute);
                row.Cells[4].SetCellValue(PipeTable[k].Attachment);
                row.Cells[5].SetCellValue(PipeTable[k].X);
                row.Cells[6].SetCellValue(PipeTable[k].Y);
                row.Cells[7].SetCellValue(PipeTable[k].H);
                row.Cells[8].SetCellValue(PipeTable[k].SPH);
                row.Cells[9].SetCellValue(PipeTable[k].EPH);
                row.Cells[10].SetCellValue(PipeTable[k].WellDepth);
                row.Cells[11].SetCellValue(PipeTable[k].SPDepth);
                row.Cells[12].SetCellValue(PipeTable[k].EPDepth);
                row.Cells[13].SetCellValue(PipeTable[k].Size);
                row.Cells[14].SetCellValue(PipeTable[k].Material);
                row.Cells[15].SetCellValue(PipeTable[k].Pressure);
                row.Cells[16].SetCellValue(PipeTable[k].Voltage);
                row.Cells[17].SetCellValue(PipeTable[k].TotalBHNum);
                row.Cells[18].SetCellValue(PipeTable[k].UsedBHNum);
                row.Cells[19].SetCellValue(PipeTable[k].CableNum);
                row.Cells[20].SetCellValue(PipeTable[k].Company);
                row.Cells[21].SetCellValue(PipeTable[k].BuryMethod);
                row.Cells[22].SetCellValue(PipeTable[k].BuryDate);
                row.Cells[23].SetCellValue(PipeTable[k].RoadName);
                row.Cells[24].SetCellValue(PipeTable[k].Comment);
            }

            using (FileStream fs = new FileStream(filePathName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        public SelectionSet SelectPipes(string selectFlag)
        {
            SelectedPoints = null;
            PromptSelectionResult psr = null;
            switch (selectFlag)
            {
                case "ALLPOINTS":
                    psr = Editor.SelectAll(new SelectionFilter(FilterInsert));
                    break;
                case "USERSELECT":
                    psr = Editor.GetSelection(new SelectionFilter(FilterInsert));
                    break;
                default:
                    break;
            }
            if (psr.Status == PromptStatus.OK)
                SelectedPoints = psr.Value;
            //CADApplication.ShowAlertDialog(SelectedPoints.Count.ToString());
            return SelectedPoints;
        }

    }
}
