using System.Collections;
using System.Drawing.Printing;

namespace MiniWord
{
    public partial class Form1 : Form
    {
        public static int docCount = 0;
        public static string txtFind = "";
        public static int startIndex = 0;
        private Font printFont= new Font("����", 10f);
        private bool isMatchCase;
        private int totalLine = 0;
        private int currentLine = 0;
        private float linesPerPage = 0f;
        private string content = null;
        private string[] arrStr = null;
        private int wordsPerPage = 0;
        private bool flag = false;
        Doc frmDoc;
        FormFind frmFind;
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void �½�ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewDoc();
            ShowStatusBar();
        }
        public void ShowStatusBar()
        {
            if (docCount > 1)
            {
                try
                {
                    toolStripStatusLabel1.Text = "����" + docCount + "���ĵ�����ǰ�����ĵ���" + base.ActiveMdiChild.Text;
                    return;
                }
                catch (Exception)
                {
                    MessageBox.Show("��ǰû�д��ڼ���״̬���Ӵ���");
                    return;
                }
            }
            if (docCount == 1)
            {
                toolStripStatusLabel1.Text = "����1���ĵ�";
            }
            else
            {
                toolStripStatusLabel1.Text = "����0���ĵ�";
            }
        }
        public void SaveDoc()
        {
            if (docCount <= 0 || base.ActiveMdiChild == null)
            {
                MessageBox.Show("�޴���");
                return;
            }
            try
            {
                frmDoc = GetActiveForm();
                if (frmDoc.firstSave)
                {
                    saveFileDialog1.FileName = frmDoc.Text;
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        frmDoc.SaveFile(saveFileDialog1.FileName);
                        frmDoc.Text = saveFileDialog1.FileName;
                        frmDoc.firstSave = false;
                        frmDoc.isSaved = true;
                    }
                }
                else
                {
                    frmDoc.SaveFile(frmDoc.Text);
                    frmDoc.isSaved = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�����ĵ�����");
            }
        }
        private Doc GetActiveForm()
        {
            try
            {
                return (Doc)base.ActiveMdiChild;
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰû�д��ڼ���״̬���Ӵ���");
                return null;
            }
        }
        private void NewDoc()
        {
            try
            {
                Doc formChild = new Doc();
                docCount++;
                formChild.Text = "�½��ĵ�" + docCount;
                formChild.MdiParent = this;
                formChild.Show();
            }
            catch (Exception)
            {
                MessageBox.Show("�½��ĵ�����");
            }
        }
        private void OpenDoc()
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Doc formChild = new Doc();
                    docCount++;
                    formChild.Text = openFileDialog1.FileName;
                    formChild.MdiParent = this;
                    formChild.LoadFile(openFileDialog1.FileName);
                    formChild.Show();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("���ĵ�����");
            }
        }
        private void CloseDoc()
        {
            if (docCount <= 0 || base.ActiveMdiChild == null)
            {
                MessageBox.Show("�޴���");
                return;
            }
            try
            {
                frmDoc = GetActiveForm();
                frmDoc.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("�ر��ĵ�����");
            }
        }

        private void ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenDoc();
            ShowStatusBar();
        }

        private void ����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveDoc();
            ShowStatusBar();
        }

        private void ���ΪToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (docCount <= 0 || base.ActiveMdiChild == null)
            {
                MessageBox.Show("�޴���");
                return;
            }
            try
            {
                frmDoc = GetActiveForm();
                saveFileDialog1.FileName = frmDoc.Text;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    frmDoc.SaveFile(saveFileDialog1.FileName);
                    frmDoc.Text = saveFileDialog1.FileName;
                    frmDoc.firstSave = false;
                    frmDoc.isSaved = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�����ĵ�����");
            }
            ShowStatusBar();
        }

        private void ȫ������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (docCount <= 0)
            {
                MessageBox.Show("�޴���");
                return;
            }
            try
            {
                Form[] mdiChildren = base.MdiChildren;
                foreach (object obj in mdiChildren)
                {
                    Doc formChild = (Doc)obj;
                    if (formChild.firstSave)
                    {
                        saveFileDialog1.FileName = formChild.Text;
                        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            formChild.SaveFile(saveFileDialog1.FileName);
                            formChild.Text = saveFileDialog1.FileName;
                            formChild.firstSave = false;
                            formChild.isSaved = true;
                        }
                    }
                    else
                    {
                        formChild.SaveFile(formChild.Text);
                        formChild.isSaved = true;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�����ĵ�����");
            }
        }

        private void �ر�ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            CloseDoc();
            ShowStatusBar();
        }

        private void ȫ���ر�ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (docCount <= 0)
            {
                MessageBox.Show("�޴���");
                return;
            }
            try
            {
                Form[] mdiChildren = base.MdiChildren;
                foreach (object obj in mdiChildren)
                {
                    Doc formChild = (Doc)obj;
                    formChild.Close();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�ر��ĵ�����");
            }
        }

        private void ҳ������ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                pageSetupDialog1.ShowDialog();
            }
            catch (Exception)
            {
                MessageBox.Show("ҳ�����ó���");
            }
        }

        private void ��ӡԤ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                printPreviewDialog1.ShowDialog();
            }
            catch (Exception)
            {
                MessageBox.Show("��ӡԤ������");
            }
        }

        private void ��ӡToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    printDocument1.Print();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("��ӡ����");
            }
            ShowStatusBar();
        }

        private void �˳�ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            NewDoc();
            ShowStatusBar();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

            OpenDoc();
            ShowStatusBar();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {

            SaveDoc();
            ShowStatusBar();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (docCount <= 0)
            {
                MessageBox.Show("�޴���");
                return;
            }
            try
            {
                Form[] mdiChildren = base.MdiChildren;
                foreach (object obj in mdiChildren)
                {
                    Doc formChild = (Doc)obj;
                    if (formChild.firstSave)
                    {
                        saveFileDialog1.FileName = formChild.Text;
                        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            formChild.SaveFile(saveFileDialog1.FileName);
                            formChild.Text = saveFileDialog1.FileName;
                            formChild.firstSave = false;
                            formChild.isSaved = true;
                        }
                    }
                    else
                    {
                        formChild.SaveFile(formChild.Text);
                        formChild.isSaved = true;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�����ĵ�����");
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            try
            {
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    printDocument1.Print();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("��ӡ����");
            }
            ShowStatusBar();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ����ToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            RichTextBox rchTxt;
            try
            {
                frmDoc = GetActiveForm();
                rchTxt = frmDoc.richTextBox1;
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰ����Ч�ĵ�");
                return;
            }
            try
            {
                if (fontDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (rchTxt.SelectedText != "")
                    {
                        rchTxt.SelectionFont = fontDialog1.Font;
                        rchTxt.SelectionColor = fontDialog1.Color;
                    }
                    else
                    {
                        rchTxt.Font = fontDialog1.Font;
                        rchTxt.ForeColor = fontDialog1.Color;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�����������");
            }
        }

        private void ��ɫToolStripMenuItem_Click(object sender, EventArgs e)
        {

            RichTextBox rchTxt;
            try
            {
                frmDoc = GetActiveForm();
                rchTxt = frmDoc.richTextBox1;
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰ����Ч�ĵ�");
                return;
            }
            try
            {
                if (colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (rchTxt.SelectedText != "")
                    {
                        rchTxt.SelectionColor = colorDialog1.Color;
                    }
                    else
                    {
                        rchTxt.ForeColor = colorDialog1.Color;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("������ɫ����");
            }
        }

        private void ����ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                if (!FormFind.isShow)
                {
                    frmFind = new FormFind(this);
                    frmFind.Show();
                }
                else
                {
                    frmFind.Activate();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("���ҳ���");
            }
        }

        public void Find(bool MatchCase)
        {
            try
            {
                isMatchCase = MatchCase;
                frmDoc = GetActiveForm();
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰ����Ч�ĵ�");
                return;
            }
            try
            {
                int num = ((!isMatchCase) ? frmDoc.richTextBox1.Find(txtFind, startIndex, RichTextBoxFinds.None) : frmDoc.richTextBox1.Find(txtFind, startIndex, RichTextBoxFinds.MatchCase));
                if (num >= 0)
                {
                    startIndex = num + 1;
                    frmDoc.richTextBox1.Focus();
                    frmDoc.richTextBox1.SelectionStart = num;
                    frmDoc.richTextBox1.SelectionLength = txtFind.Length;
                }
                else
                {
                    MessageBox.Show("�������");
                    startIndex = 0;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("�����ı�����");
            }
        }
        private void ������һ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Find(isMatchCase);
        }

        private void ��������ʱ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                frmDoc = GetActiveForm();
                frmDoc.richTextBox1.SelectedText = DateTime.Now.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰ����Ч�ĵ�");
            }
        }

        private void �Զ�����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            �Զ�����ToolStripMenuItem.Checked = !�Զ�����ToolStripMenuItem.Checked;
            try
            {
                frmDoc = GetActiveForm();
                frmDoc.richTextBox1.WordWrap = !frmDoc.richTextBox1.WordWrap;
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰ����Ч�ĵ�");
            }
        }

        private void ȫѡToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                frmDoc = GetActiveForm();
                frmDoc.richTextBox1.SelectAll();
            }
            catch (Exception)
            {
                MessageBox.Show("��ǰ����Ч�ĵ�");
            }
        }

        private void ���ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void ƽ��ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ����ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormAbout frmAbout = new FormAbout();
            frmAbout.ShowDialog();
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {

            Graphics graphics = e.Graphics;
            float num = 0f;
            float num2 = e.MarginBounds.Left;
            float num3 = e.MarginBounds.Top;
            int num4 = 0;
            if (!flag)
            {
                linesPerPage = (float)e.MarginBounds.Height / printFont.GetHeight(graphics);
                wordsPerPage = (int)((float)e.MarginBounds.Width / printFont.GetHeight(graphics));
                frmDoc = GetActiveForm();
                content = frmDoc.richTextBox1.Text;
                content = FormatCuttingString(content, wordsPerPage);
                arrStr = content.Split('\n');
                totalLine = arrStr.Length;
                flag = !flag;
            }
            num4 = 0;
            while ((float)num4 <= linesPerPage && currentLine < totalLine)
            {
                num = num3 + (float)num4 * printFont.GetHeight(graphics);
                graphics.DrawString(arrStr[currentLine], printFont, Brushes.Black, num2, num, new StringFormat());
                num4++;
                currentLine++;
            }
            if (currentLine < totalLine)
            {
                e.HasMorePages = true;
                return;
            }
            e.HasMorePages = false;
            flag = false;
            currentLine = 0;
        }

        private string FormatCuttingString(string s, int n)
        {
            int num = 0;
            ArrayList arrayList = new ArrayList();
            string text = "";
            if (s.Length <= n)
            {
                return s;
            }
            for (int i = 0; i < s.Length; i++)
            {
                num++;
                arrayList.Add(s[i]);
                if (s[i] == '\n')
                {
                    num = 0;
                }
                if (num == n)
                {
                    arrayList.Add('\n');
                    num = 0;
                }
            }
            foreach (object item in arrayList)
            {
                char c = (char)item;
                text += c;
            }
            return text;
        }
    }

}