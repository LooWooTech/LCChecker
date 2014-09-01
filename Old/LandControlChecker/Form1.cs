using LooWoo.Land.LandControlChecker;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ImportUtil
{
    public partial class Form1 : Form
    {
        
        private readonly FolderBrowserDialog folderDialog = new FolderBrowserDialog()
            {
                ShowNewFolderButton = false,
                Description = "请选择需要检查的表格文件所在目录"
            };

        private readonly OpenFileDialog fileDialg = new OpenFileDialog()
            {
                AddExtension = true,
                CheckFileExists = false,
                Filter = "Excel 97-2003文件 (*.xls) |*.xls|Excel 2007文件 (*.xlsx)|*.xlsx",
                Title = "请选择检查结果表格文件"
            };

        public Form1()
        {
            InitializeComponent();
        }

        private void btnDir_Click(object sender, EventArgs e)
        {
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = folderDialog.SelectedPath;
            }
        }

        private void btnFile_Click(object sender, EventArgs e)
        {
            if (fileDialg.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = fileDialg.FileName;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFolder.Text) == true)
            {
                MessageBox.Show("请选择导入目录。");
                return;
            }
            if (string.IsNullOrEmpty(txtFile.Text) == true)
            {
                MessageBox.Show("请选择导出文件。");
                return;
            }

            btnDir.Enabled = false;
            btnFile.Enabled = false;
            btnExport.Enabled = false;
            btnImport.Enabled = false;

            var engine = new CheckEngine();
            engine.CheckStatusChanged += CheckEngine_CheckStatusChanged;

            //try
            {
                listViewDict.Clear();
                listView1.Items.Clear();
                
                var ret = engine.Check(txtFolder.Text, txtFile.Text, radioButton1.Checked);
                MessageBox.Show(string.Format("已成功完成对{0}个文件的检查。", ret), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //catch (Exception ex)
            {
                //MessageBox.Show("检查时发生错误：\r\n" + ex, "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
               
            }

            btnDir.Enabled = true;
            btnFile.Enabled = true;
            btnExport.Enabled = true;
            btnImport.Enabled = true;
        }

        private Dictionary<string, ListViewItem> listViewDict = new Dictionary<string, ListViewItem>(); 

        private void CheckEngine_CheckStatusChanged(object sender, CheckStatusChangedEventArgs e)
        {
            if (listViewDict.ContainsKey(e.File))
            {
                listViewDict[e.File].SubItems[1].Text = e.Text;
            }
            else
            {
                var item = new ListViewItem(e.File);
                item.SubItems.Add(e.File);
                item.SubItems.Add(e.Text);
                listViewDict.Add(e.File, item);
                listView1.Items.Add(item);
            }
            Application.DoEvents();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            var dialog = new SaveFileDialog()
                {
                    Title = "请选择导出文件",
                    Filter = "CSV文件 (*.csv)|*.csv"
                };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var stream = new StreamWriter(dialog.FileName, false, Encoding.GetEncoding("GB2312")))
                    {
                        for (var i = 0; i < listView1.Items.Count; i++)
                        {
                            stream.WriteLine("{0},{1}", listView1.Items[i].SubItems[0].Text, listView1.Items[i].SubItems[1].Text);
                        }
                        MessageBox.Show("导出完成。");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出时发生错误:" + ex.Message);
                }
            }

        }
    }
}
