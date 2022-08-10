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
using EXLApplication = NetOffice.ExcelApi.Application;
using NetOffice.ExcelApi;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

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
            ISheet sheet = null;
            FileStream fs;
            using (fs = File.Open(filePathName, FileMode.Open, FileAccess.Read))
            {
                // 2007版本
                if (filePathName.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                // 2003版本
                else if (filePathName.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                sheet = workbook.GetSheetAt(0);
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum + 1;
                    if (rowCount > 1)
                    {
                        for (ushort irow = 1; irow < sheet.LastRowNum; irow++)
                        {
                            IRow row = sheet.GetRow(irow);

                            var pipe = new PipeLineProperty
                            {
                                RowInd = irow,
                                Name = row.GetCell(0) == null ? "" : row.GetCell(0).StringCellValue,
                                WTName = row.GetCell(1) == null ? "" : row.GetCell(1).StringCellValue,
                                Connect = row.GetCell(2) == null ? "" : row.GetCell(2).StringCellValue,
                                Attribute = row.GetCell(3) == null ? "" : row.GetCell(3).StringCellValue,
                                Attachment = row.GetCell(4) == null ? "" : row.GetCell(4).StringCellValue,
                                X = row.GetCell(5) == null ? 0 : row.GetCell(5).NumericCellValue,
                                Y = row.GetCell(6) == null ? 0 : row.GetCell(6).NumericCellValue,
                                H = row.GetCell(7) == null ? 0 : row.GetCell(7).NumericCellValue,
                                SPH = row.GetCell(8) == null ? 0 : row.GetCell(8).NumericCellValue,
                                EPH = row.GetCell(9) == null ? 0 : row.GetCell(9).NumericCellValue,
                                WellDepth = row.GetCell(10) == null ? 0 : row.GetCell(10).NumericCellValue,
                                SPDepth = row.GetCell(11) == null ? 0 : row.GetCell(11).NumericCellValue,
                                EPDepth = row.GetCell(12) == null ? 0 : row.GetCell(12).NumericCellValue,
                                Size = row.GetCell(13) == null ? "" : row.GetCell(14).StringCellValue,
                                Material = row.GetCell(14) == null ? "" : row.GetCell(14).StringCellValue,
                                Pressure = row.GetCell(15) == null ? "" : row.GetCell(15).StringCellValue,
                                Voltage = row.GetCell(16) == null ? "" : row.GetCell(16).StringCellValue,
                                TotalBHNum = row.GetCell(17) == null ? (ushort)0 : (ushort)row.GetCell(17).NumericCellValue,
                                UsedBHNum = row.GetCell(18) == null ? (ushort)0 : (ushort)row.GetCell(18).NumericCellValue,
                                CableNum = row.GetCell(19) == null ? (ushort)0 : (ushort)row.GetCell(19).NumericCellValue,
                                Company = row.GetCell(20) == null ? "" : row.GetCell(20).StringCellValue,
                                BuryMethod = row.GetCell(21) == null ? "" : row.GetCell(21).StringCellValue,
                                BuryDate = row.GetCell(22) == null ? "" : row.GetCell(22).StringCellValue,
                                RoadName = row.GetCell(23) == null ? "" : row.GetCell(23).StringCellValue,
                                Comment = row.GetCell(24) == null ? "" : row.GetCell(24).StringCellValue
                            };
                            pipe.PipeLineType = (PipeLineType)Enum.Parse(typeof(PipeLineType), pipe.WTName.Substring(0, 2));
                            Pipes.Add(pipe);
                        }
                        ed.WriteMessage("ok");
                    }
                }
            }
        }
    }
}
