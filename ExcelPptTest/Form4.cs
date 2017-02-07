using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPptTest
{
    public partial class Form4 : Form
    {
        System.Drawing.Imaging.Metafile img = null;

        public Form4()
        {
            InitializeComponent();
            //double d1 = 20;
            //double d2 = d1 / 100;

            img = new System.Drawing.Imaging.Metafile(@"D:\test\Test Slide\Slide1.EMF");
        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.DrawImage(img, new Rectangle(Point.Empty, pictureBox1.ClientSize));
        }

        private void pictureBox1_Resize(object sender, EventArgs e)
        {
            pictureBox1.Invalidate();
        }
    }
}
