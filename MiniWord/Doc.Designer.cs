namespace MiniWord
{
    partial class Doc
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 12);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(547, 359);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            // 
            // Doc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(571, 383);
            this.Controls.Add(this.richTextBox1);
            this.Name = "Doc";
            this.Text = "Doc";
            this.Activated += new System.EventHandler(this.Doc_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Doc_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Doc_FormClosed);
            this.Load += new System.EventHandler(this.Doc_Load);
            this.ResumeLayout(false);

        }

        #endregion

        public RichTextBox richTextBox1;
    }
}