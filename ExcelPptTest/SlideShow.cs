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
    public partial class SlideShow : Form
    {
        private int imageNumber = 1;

        public SlideShow()
        {
            InitializeComponent();
            tblBusy.Visible = false;
            //tblSlideContainer.Controls.Add(tblBusy);

            //tblBusy.Top = picSlides.Height - tblBusy.Height / 2;
            //tblBusy.Left = picSlides.Width - tblBusy.Width / 2;

            //tblSlideContainer.Controls.Add(tblBusy, 0, 1);
            picSlides.Controls.Add(tblBusy);
            tblBusy.Visible = true;
            tblBusy.BringToFront();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //timer1.Enabled = true;
            //int BorderWidth = (this.Width - this.ClientSize.Width) /2;
            //int TitlebarHeight = this.Height - this.ClientSize.Height - 2 * BorderWidth;
            //pictureBox1.Height = this.Height - TitlebarHeight - panel1.Height;

            //var imageSize = pictureBox1.Image.Size;
            //var fitSize = pictureBox1.ClientSize;
            //pictureBox1.SizeMode = imageSize.Width > fitSize.Width || imageSize.Height > fitSize.Height ?
            //    PictureBoxSizeMode.Zoom : PictureBoxSizeMode.CenterImage;

            tblBusy.Visible = true;
            loadNextImage();
            timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            loadNextImage();
        }

        private void loadNextImage()
        {
            tblBusy.Visible = false;
            if (imageNumber == 5)
            {
                imageNumber = 1;
            }
            picSlides.ImageLocation = string.Format(@"E:\VSTS\Source\WorkSpaces\NPSData\NPSData\bin\Debug\Slides\Slide{0}.EMF", imageNumber);
            imageNumber++;
        }


    }
}
