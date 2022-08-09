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

namespace AutoPipelines
{
    public class Main
    {
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

        private void ReadPropertyTab(string fileName)
        {
            Editor ed = CADApplication.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(fileName);

            EXLApplication excelApp = new EXLApplication();
            Workbook workbook = excelApp.Workbooks.Open(fileName);
            //_excelApp.Visible = true; 
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet;

            var rowsCount = worksheet.UsedRange.Rows.Count();
            List<PipeLineProperty> pipes = new List<PipeLineProperty>();
            pipes.Init(rowsCount - 1);
                for (ushort irow = 2; irow <= rowsCount; irow++)
                {
                    var ipipe = new PipeLineProperty
                    {
                        RowInd = irow,
                        Name = worksheet.Cells[irow, 1].Value == null ? "" : worksheet.Cells[irow, 1].Value.ToString(),
                        WTName = worksheet.Cells[irow, 2].Value == null ? "" : worksheet.Cells[irow, 2].Value.ToString()
                    };
                    ipipe.PipeLineType = (PipeLineType)Enum.Parse(typeof(PipeLineType), ipipe.WTName.Substring(0, 2));
                    ipipe.Connect = worksheet.Cells[irow, 3].Value == null ? "" : worksheet.Cells[irow, 3].Value.ToString();
                    ipipe.Attribute = worksheet.Cells[irow, 4].Value == null ? "" : worksheet.Cells[irow, 4].Value.ToString();
                    ipipe.Attachment = worksheet.Cells[irow, 5].Value == null ? "" : worksheet.Cells[irow, 5].Value.ToString();
                    ipipe.X = worksheet.Cells[irow, 6].Value == null ? 0 : (double)worksheet.Cells[irow, 6].Value;
                    ipipe.Y = worksheet.Cells[irow, 7].Value == null ? 0 : (double)worksheet.Cells[irow, 7].Value;
                    ipipe.H = worksheet.Cells[irow, 8].Value == null ? 0 : (double)worksheet.Cells[irow, 8].Value;
                    ipipe.SPH = worksheet.Cells[irow, 9].Value == null ? 0 : (double)worksheet.Cells[irow, 9].Value;
                    ipipe.EPH = worksheet.Cells[irow, 10].Value == null ? 0 : (double)worksheet.Cells[irow, 10].Value;
                    ipipe.WellDepth = worksheet.Cells[irow, 11].Value == null ? 0 : (double)worksheet.Cells[irow, 11].Value;
                    ipipe.SPDepth = worksheet.Cells[irow, 12].Value == null ? 0 : (double)worksheet.Cells[irow, 12].Value;
                    ipipe.EPDepth = worksheet.Cells[irow, 13].Value == null ? 0 : (double)worksheet.Cells[irow, 13].Value;
                    ipipe.Size = worksheet.Cells[irow, 14].Value == null ? "" : worksheet.Cells[irow, 14].Value.ToString();
                    ipipe.Material = worksheet.Cells[irow, 15].Value == null ? "" : worksheet.Cells[irow, 15].Value.ToString();
                    ipipe.Pressure = worksheet.Cells[irow, 16].Value == null ? "" : worksheet.Cells[irow, 16].Value.ToString();
                    ipipe.Voltage = worksheet.Cells[irow, 17].Value == null ? "" : worksheet.Cells[irow, 17].Value.ToString();
                    ipipe.TotalBHNum = (ushort)(worksheet.Cells[irow, 18].Value == null ? 0 : (ushort)worksheet.Cells[irow, 18].Value);
                    ipipe.UsedBHNum = (ushort)(worksheet.Cells[irow, 19].Value == null ? 0 : (ushort)worksheet.Cells[irow, 19].Value);
                    ipipe.CableNum = (ushort)(worksheet.Cells[irow, 20].Value == null ? 0 : (ushort)worksheet.Cells[irow, 20].Value);
                    ipipe.Company = worksheet.Cells[irow, 21].Value == null ? "" : worksheet.Cells[irow, 21].Value.ToString();
                    ipipe.BuryMethod = worksheet.Cells[irow, 22].Value == null ? "" : worksheet.Cells[irow, 22].Value.ToString();
                    ipipe.BuryDate = worksheet.Cells[irow, 23].Value == null ? "" : worksheet.Cells[irow, 23].Value.ToString();
                    ipipe.RoadName = worksheet.Cells[irow, 24].Value == null ? "" : worksheet.Cells[irow, 24].Value.ToString();
                    ipipe.Comment = worksheet.Cells[irow, 25].Value == null ? "" : worksheet.Cells[irow, 25].Value.ToString();
                    pipes.Add(ipipe);
                }
                Console.WriteLine(pipes);
                workbook.Dispose();
                excelApp.Dispose();
            //ed.WriteMessage(worksheet.Cells[irow, 1].Value.ToString());
        }
    }
    }
