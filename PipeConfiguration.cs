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
    public partial class PipeConfiguration : Form
    {
        AutoPipe AutoPipe;

        public PipeConfiguration()
        {
            InitializeComponent();
            pipeNamePosCmbBox.SelectedIndex = 1;
            lineStyleCmbBox.SelectedIndex = 0;
            fzlPosCmbBox.SelectedIndex = 1;
            toolStripStatusLabel1.Text = "lq5991@csepdi.com";
        }

        private void openTabBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Title = "打开属性表",
                Filter = "Excel工作簿(*.xls,*.xlsx)|*.xls;*.xlsx",
                InitialDirectory = @"C:\Users\Channing\source\repos\LiQian-2019\AutoPipelines\PropertyTabs\"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                tabFileTxtBox.Text = ofd.FileName;
                executeDrawBtn.Enabled = false;
                AutoPipe = new AutoPipe() { TabFilePathName = tabFileTxtBox.Text };
                AutoPipe.ReadPropertyTab();
                toolStripStatusLabel1.Text = "属性表读取完成。";
            }

        }

        private void cancelConfigBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void executeDrawBtn_Click(object sender, EventArgs e)
        {
            if (AutoPipe.ErrPipeInd.Any())
            {
                MessageBoxButtons box = MessageBoxButtons.OKCancel;
                DialogResult result = MessageBox.Show("属性表中有错误，是否强行绘图？", "提示", box);
                if (result == DialogResult.OK)
                {
                    foreach (var errorPipeInd in AutoPipe.ErrPipeInd)
                    {
                        switch (errorPipeInd.Value)
                        {
                            case 3:
                                var errorPipe = AutoPipe.PipeTable.Find(p => p.Name == errorPipeInd.Key);
                                errorPipe.Attachment = "";
                                errorPipe.Attribute = "一般管线点";
                                break;
                            default:
                                errorPipe = AutoPipe.PipeTable.Find(p => p.Name == errorPipeInd.Key);
                                AutoPipe.PipeTable.Remove(errorPipe);
                                break;
                        }
                    }
                }
                else return;
            }
            AutoPipe.InputConfigPara(this);
            AutoPipe.DrawPipes();
            toolStripStatusLabel1.Text = "管线图绘制完毕。";
            toolStripProgressBar1.Value = 0;
        }

        private void checkTabBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tabFileTxtBox.Text))
                MessageBox.Show("请输入属性表路径！");
            else
            {
                AutoPipe.AddRowValue += new RowAdd(GridViewAddRow);
                AutoPipe.AddBarValue += new BarGrow(ProgressBarGrow);
                toolStripProgressBar1.Value = 0;
                //ClearAllRows(checkResultDataGrid);
                checkResultDataGrid.Rows.Clear();
                AutoPipe.CheckPropertyTab();
                toolStripProgressBar1.Value = 0;

                executeDrawBtn.Enabled = true;
            }
        }

        private void drawPipeFzlChkBox_MouseClick(object sender, MouseEventArgs e)
        {
            var cb = (CheckBox)sender;
            if (!cb.Checked)
                foreach (var item in groupBox4.Controls)
                    ((Control)item).Enabled = false;
            else
                foreach (var item in groupBox4.Controls)
                    ((Control)item).Enabled = true;
        }

        private void drawPipeNameChkBox_MouseClick(object sender, MouseEventArgs e)
        {
            var cb = (CheckBox)sender;
            if (!cb.Checked)
                foreach (var item in groupBox1.Controls)
                    ((Control)item).Enabled = false;
            else
                foreach (var item in groupBox1.Controls)
                    ((Control)item).Enabled = true;
        }

        private void drawPipeLineChkBox_MouseClick(object sender, MouseEventArgs e)
        {
            var cb = (CheckBox)sender;
            if (!cb.Checked)
            {
                foreach (var item in groupBox2.Controls)
                    ((Control)item).Enabled = false;
                foreach (var item in groupBox3.Controls)
                {
                    ((CheckBox)item).Checked = false;
                    ((CheckBox)item).Enabled = false;
                }
            }
            else
            {
                foreach (var item in groupBox2.Controls)
                    ((Control)item).Enabled = true;
                foreach (var item in groupBox3.Controls)
                {
                    ((CheckBox)item).Checked = true;
                    ((CheckBox)item).Enabled = true;
                }
            }
        }

        public void GridViewAddRow(string[] rowContent)
        {
            this.checkResultDataGrid.Rows.Add(rowContent);
        }

        public void ProgressBarGrow(int value)
        {
            this.toolStripProgressBar1.Value = value;
        }

        public void ClearAllRows(DataGridView gridView)
        {
            if (gridView.RowCount < 1) return;
            int i = 0;
            while (i >= 0)
            {
                i = gridView.RowCount - 1;
                gridView.Rows.Remove(gridView.Rows[i]);
                i--;
            }
        }
    }
}
