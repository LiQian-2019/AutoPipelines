using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoPipelines
{
    public partial class UCRecordPM : UserControl
    {
        public string CurrentHandle { get; set; }
        public UCRecordPM()
        {
            InitializeComponent();
            RecorPMInit();
            CurrentHandle = string.Empty;
        }

        public void RecorPMInit()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add(25);
            this.dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dataGridView1.Rows[0].Cells[0].Value = "所在行数";
            this.dataGridView1.Rows[1].Cells[0].Value = "图上点号";
            this.dataGridView1.Rows[2].Cells[0].Value = "物探点号";
            this.dataGridView1.Rows[3].Cells[0].Value = "连接点号";
            this.dataGridView1.Rows[4].Cells[0].Value = "特征点";
            this.dataGridView1.Rows[5].Cells[0].Value = "附属物名称";
            this.dataGridView1.Rows[6].Cells[0].Value = "X坐标(m)";
            this.dataGridView1.Rows[7].Cells[0].Value = "Y坐标(m)";
            this.dataGridView1.Rows[8].Cells[0].Value = "Z坐标(m)";
            this.dataGridView1.Rows[9].Cells[0].Value = "起点高程(m)";
            this.dataGridView1.Rows[10].Cells[0].Value = "终点高程(m)";
            this.dataGridView1.Rows[11].Cells[0].Value = "井深(m)";
            this.dataGridView1.Rows[12].Cells[0].Value = "起点埋深(m)";
            this.dataGridView1.Rows[13].Cells[0].Value = "终点埋深(m)";
            this.dataGridView1.Rows[14].Cells[0].Value = "管径尺寸";
            this.dataGridView1.Rows[15].Cells[0].Value = "材质";
            this.dataGridView1.Rows[16].Cells[0].Value = "压力";
            this.dataGridView1.Rows[17].Cells[0].Value = "电压";
            this.dataGridView1.Rows[18].Cells[0].Value = "总孔数";
            this.dataGridView1.Rows[19].Cells[0].Value = "已用孔数";
            this.dataGridView1.Rows[20].Cells[0].Value = "电缆条数";
            this.dataGridView1.Rows[21].Cells[0].Value = "权属单位";
            this.dataGridView1.Rows[22].Cells[0].Value = "埋设方式";
            this.dataGridView1.Rows[23].Cells[0].Value = "埋设日期";
            this.dataGridView1.Rows[24].Cells[0].Value = "道路名称";
            this.dataGridView1.Rows[25].Cells[0].Value = "备注";
        }

        public void UpdatePM(string handle, TypedValue[] values)
        {
            this.dataGridView1.Rows[0].Cells[1].Value = values[0].Value;
            this.dataGridView1.Rows[1].Cells[1].Value = values[1].Value;
            this.dataGridView1.Rows[2].Cells[1].Value = values[2].Value;
            //this.dataGridView1.Rows[3].Cells[1].Value = values[3].Value.ToString();
            this.dataGridView1.Rows[3].Cells[1].Value = values[4].Value;
            this.dataGridView1.Rows[4].Cells[1].Value = values[5].Value;
            this.dataGridView1.Rows[5].Cells[1].Value = values[6].Value;
            this.dataGridView1.Rows[6].Cells[1].Value = values[7].Value;
            this.dataGridView1.Rows[7].Cells[1].Value = values[8].Value;
            this.dataGridView1.Rows[8].Cells[1].Value = values[9].Value;
            this.dataGridView1.Rows[9].Cells[1].Value = values[10].Value;
            this.dataGridView1.Rows[10].Cells[1].Value = values[11].Value;
            this.dataGridView1.Rows[11].Cells[1].Value = values[12].Value;
            this.dataGridView1.Rows[12].Cells[1].Value = values[13].Value;
            this.dataGridView1.Rows[13].Cells[1].Value = values[14].Value;
            this.dataGridView1.Rows[14].Cells[1].Value = values[15].Value;
            this.dataGridView1.Rows[15].Cells[1].Value = values[16].Value;
            this.dataGridView1.Rows[16].Cells[1].Value = values[17].Value;
            this.dataGridView1.Rows[17].Cells[1].Value = values[18].Value;
            this.dataGridView1.Rows[18].Cells[1].Value = values[19].Value;
            this.dataGridView1.Rows[19].Cells[1].Value = values[20].Value;
            this.dataGridView1.Rows[20].Cells[1].Value = values[21].Value;
            this.dataGridView1.Rows[21].Cells[1].Value = values[22].Value;
            this.dataGridView1.Rows[22].Cells[1].Value = values[23].Value;
            this.dataGridView1.Rows[23].Cells[1].Value = values[24].Value;
            this.dataGridView1.Rows[24].Cells[1].Value = values[25].Value;
            this.dataGridView1.Rows[25].Cells[1].Value = values[26].Value;
            CurrentHandle = handle;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            MessageBox.Show($"第{e.RowIndex}行，第{e.ColumnIndex}列发生了修改！");
        }

        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            MessageBox.Show("Test");
            // 执行检查数值合法性的方法
        }


    }
}
