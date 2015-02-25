using CoordAnalyseService;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CoordAnalyseClient
{
    public partial class Form1 : Form
    {
        private readonly FolderBrowserDialog folderDialog = new FolderBrowserDialog()
        {
            ShowNewFolderButton = false,
            Description = "请选择坐标文件所在目录"
        };

        private readonly FileDialog fileDialg = new OpenFileDialog()
        {
            AddExtension = true,
            CheckFileExists = false,
            Filter = "zip压缩包文件 (*.zip) |*.zip",
            Title = "请选择文件"
        };

        public Form1()
        {
            InitializeComponent();
            Analyser.ShowInfoHandler = ShowInfo;
        }

        private void btnDir_Click(object sender, EventArgs e)
        {
            
            if (fileDialg.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = fileDialg.FileName;
            }
        }

        private void btnFile_Click(object sender, EventArgs e)
        {
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = folderDialog.SelectedPath;
            }
        }

        private void ShowInfo(string projectNo, string msg)
        {
            var item = new ListViewItem(new[] { projectNo, msg });
            listView1.Items.Add(item);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            Analyser.ProcessNext(txtFolder.Text, txtFile.Text, true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            Analyser.ProcessNext(txtFolder.Text, txtFile.Text, false);
        }
    }
}
