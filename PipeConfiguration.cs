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
            }

        }

        private void cancelConfigBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void executeDrawBtn_Click(object sender, EventArgs e)
        {
            AutoPipe.DrawPipes();
            toolStripStatusLabel1.Text = "管线图绘制完毕。";
        }

        private void checkTabBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tabFileTxtBox.Text))
                MessageBox.Show("请输入属性表路径！");
            else
            {
                AutoPipe = new AutoPipe(this) { TabFilePathName = tabFileTxtBox.Text };
                AutoPipe.ReadPropertyTab();
                toolStripProgressBar1.Value = 50;
                AutoPipe.CheckPropertyTab();
                toolStripProgressBar1.Value = 100;

                executeDrawBtn.Enabled = true;
                toolStripProgressBar1.Value = 0;
            }
        }

        private void drawPipeTextChkBox_MouseClick(object sender, MouseEventArgs e)
        {

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
    }
}
