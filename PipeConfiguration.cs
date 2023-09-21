using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoPipelines
{
    public partial class PipeConfiguration : Form
    {
        readonly AutoPipe AutoPipe;

        public PipeConfiguration()
        {
            InitializeComponent();
            pipeNamePosCmbBox.SelectedIndex = 1;
            lineStyleCmbBox.SelectedIndex = 0;
            fzlPosCmbBox.SelectedIndex = 1;
            toolStripStatusLabel1.Text = "lq5991@csepdi.com";
            AutoPipe = new AutoPipe();
        }

        private void openTabBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Title = "打开属性表",
                Filter = "Excel工作簿(*.xls,*.xlsx)|*.xls;*.xlsx",
                InitialDirectory = System.Environment.CurrentDirectory
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                tabFileTxtBox.Text = ofd.FileName;
                AutoPipe.TabFilePathName = tabFileTxtBox.Text;
                AutoPipe.ReadPropertyTab();
                toolStripStatusLabel1.Text = "属性表读取完成。";
                checkTabBtn.Enabled = true;
                executeDrawBtn.Enabled = false;
            }

        }

        private void cancelConfigBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void executeDrawBtn_Click(object sender, EventArgs e)
        {
            if (AutoPipe.ErrPipeInf.Any())
            {
                MessageBoxButtons box = MessageBoxButtons.OKCancel;
                DialogResult result = MessageBox.Show("属性表中有错误，是否强行绘图？", "提示", box);
                if (result == DialogResult.OK)
                {
                    foreach (var errPipeInf in AutoPipe.ErrPipeInf)
                    {
                        var errorPipe = AutoPipe.RawPipeTable.Find(p => p.Name == errPipeInf[1]);
                        if (errPipeInf[3].EndsWith("块文件") || errPipeInf[3].EndsWith("名称"))
                        {
                            errorPipe.Attachment = "";
                            errorPipe.Attribute = "一般管线点";
                            string blockName = errorPipe.PipeLineType + "P一般管线点";
                            if (!AutoPipe.CadBlockTable.Has(blockName))
                                AutoPipe.InsertCADBlock(blockName);
                        }
                        else if (errPipeInf[3] == "未找到与之相连的点号")
                        {
                            errorPipe.Connect = "";
                        }
                        else
                        {
                            AutoPipe.RawPipeTable.Remove(errorPipe);
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
                staticBtn.Enabled = true;
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

        private void allPipeTypesRdoBtn_CheckedChanged(object sender, EventArgs e)
        {
            this.selfDefinedTextBox.Enabled = false;
        }

        private void customPipeTypesRdoBtn_CheckedChanged(object sender, EventArgs e)
        {
            this.selfDefinedTextBox.Enabled = true;
        }

        private void staticBtn_Click(object sender, EventArgs e)
        {
            string text = string.Empty;
            foreach (var pll in AutoPipe.PipeLineLength)
            {
                text += pll.Key + "的总长度为：" + pll.Value.ToString("0.00") + "m\n";
            }
            text += "所有管段总长度为：" + AutoPipe.PipeLineLength.Sum(x => x.Value).ToString("F2") + "m";
            DialogResult d = MessageBox.Show(text, "统计（点击『确定』复制到剪贴板）", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (d == DialogResult.OK)
            {
                Clipboard.SetText(text);
            }
        }
    }
}
