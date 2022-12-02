using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniWord
{
    public partial class Doc : Form
    {
        public bool isSaved = false;
        public bool firstSave = true;
        public Form1 frmParent;
        public Doc()
        {
            InitializeComponent();
        }

        private void Doc_Load(object sender, EventArgs e)
        {
            richTextBox1.Dock = DockStyle.Fill;
            richTextBox1.WordWrap = false;
            richTextBox1.DetectUrls = true;
        }
        public void LoadFile(string filename)
        {
            try
            {
                richTextBox1.LoadFile(filename);
                isSaved = true;
                firstSave = false;
            }
            catch (Exception)
            {
                MessageBox.Show("文件打开错误");
            }
        }

        public void SaveFile(string filename)
        {
            try
            {
                richTextBox1.SaveFile(filename);
                isSaved = true;
                firstSave = false;
            }
            catch (Exception)
            {
                MessageBox.Show("文件保存错误");
            }
        }

        private void Doc_Activated(object sender, EventArgs e)
        {
            RefreshParentStatursBar();
        }
        private void RefreshParentStatursBar()
        {
            frmParent = GetParentForm();
            frmParent.ShowStatusBar();
        }
        private Form1 GetParentForm()
        {
            try
            {
                return (Form1)base.MdiParent;
            }
            catch (Exception)
            {
                MessageBox.Show("父窗体错误");
                return null;
            }
        }

        private void Doc_FormClosed(object sender, FormClosedEventArgs e)
        {

            Form1.docCount--;
            RefreshParentStatursBar();
        }

        private void Doc_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!isSaved)
            {
                switch (MessageBox.Show("是否要保存文档?\n\n" + Text, "关闭", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
                {
                    case DialogResult.Yes:
                        frmParent = GetParentForm();
                        frmParent.SaveDoc();
                        RefreshParentStatursBar();
                        break;
                    case DialogResult.Cancel:
                        e.Cancel = true;
                        break;
                }
            }
        }
    }
}
