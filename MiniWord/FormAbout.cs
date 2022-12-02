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
    public partial class FormAbout : Form
    {
        
        private Random rnd = new Random();
        public FormAbout()
        {
            InitializeComponent();
            pictureBox1.Image = Image.FromFile("C:\\Users\\Jan29th\\source\\repos\\MiniWord\\MiniWord\\pic\\0.jpg");
            timer1.Start();
        }

        private void FormAbout_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            int num = rnd.Next(0, 7);
            try
            {
                pictureBox1.Image = Image.FromFile("C:\\Users\\Jan29th\\source\\repos\\MiniWord\\MiniWord\\pic\\" + num + ".jpg");
            }
            catch (Exception)
            {
                return;
            }
        }
    }
}
