using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoPipelines
{
    public partial class ExportToXlsx : Form
    {
        public ExportTable ExpTab { get; set; }
        public ExportToXlsx()
        {
            InitializeComponent();
            ExpTab = new ExportTable();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                Title = "选择属性表导出位置",
                Filter = "Excel工作簿(*.xlsx)|*.xlsx",
                InitialDirectory = @"C:\Users\Channing\source\repos\LiQian-2019\AutoPipelines\PropertyTabs\"
            };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = sfd.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportButton();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Enabled = !textBox2.Enabled;
        }

        private void ExportButton()
        {            
            if(textBox1.Text == "")
            {
                _ = MessageBox.Show("导出路径为空！请检查路径。", "错误", 0);
                return;
            }
            ExpTab.FilePathName = textBox1.Text;

            if (radioButton1.Checked)
                ExpTab.SelectedPoints = ExpTab.SelectPipes("ALLPOINTS");

            if (radioButton3.Checked)
                ExpTab.PipePropStr = "ALLTYPES";
            else if (radioButton4.Checked)
            {
                if (textBox2.Text == "")
                {
                    _ = MessageBox.Show("未输入任何管类！", "错误", 0);
                    return;
                }
                ExpTab.PipePropStr = textBox2.Text;
            }

            if(ExpTab.SelectedPoints == null)
            {
                MessageBox.Show("未选择任何管点！","错误",0);
                return;
            }
            if (ExpTab.ExcuteExport(ExpTab.SelectedPoints))
                MessageBox.Show("属性表导出完成。","提示",0);
        }

        private void button2_KeyPress(object sender, KeyPressEventArgs e)
        {
            ExportButton();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            button4.Enabled = !button4.Enabled;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExpTab.SelectedPoints = ExpTab.SelectPipes("USERSELECT");
        }
    }
}
