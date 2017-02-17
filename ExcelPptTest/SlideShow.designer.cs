namespace ExcelPptTest
{
    partial class SlideShow
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
            this.components = new System.ComponentModel.Container();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.tblSlideContainer = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.picSlides = new System.Windows.Forms.PictureBox();
            this.tblBusy = new System.Windows.Forms.TableLayoutPanel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.tblSlideContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picSlides)).BeginInit();
            this.tblBusy.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Interval = 5000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // tblSlideContainer
            // 
            this.tblSlideContainer.ColumnCount = 1;
            this.tblSlideContainer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblSlideContainer.Controls.Add(this.tblBusy, 0, 2);
            this.tblSlideContainer.Controls.Add(this.panel1, 0, 1);
            this.tblSlideContainer.Controls.Add(this.picSlides, 0, 0);
            this.tblSlideContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblSlideContainer.Location = new System.Drawing.Point(0, 0);
            this.tblSlideContainer.Name = "tblSlideContainer";
            this.tblSlideContainer.RowCount = 3;
            this.tblSlideContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 95F));
            this.tblSlideContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblSlideContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tblSlideContainer.Size = new System.Drawing.Size(949, 454);
            this.tblSlideContainer.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.DodgerBlue;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 415);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(943, 15);
            this.panel1.TabIndex = 2;
            // 
            // picSlides
            // 
            this.picSlides.BackColor = System.Drawing.Color.White;
            this.picSlides.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.picSlides.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picSlides.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picSlides.Location = new System.Drawing.Point(3, 3);
            this.picSlides.Name = "picSlides";
            this.picSlides.Size = new System.Drawing.Size(943, 406);
            this.picSlides.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picSlides.TabIndex = 1;
            this.picSlides.TabStop = false;
            // 
            // tblBusy
            // 
            this.tblBusy.BackColor = System.Drawing.Color.White;
            this.tblBusy.ColumnCount = 1;
            this.tblBusy.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblBusy.Controls.Add(this.panel2, 0, 1);
            this.tblBusy.Controls.Add(this.pictureBox2, 0, 0);
            this.tblBusy.Location = new System.Drawing.Point(3, 436);
            this.tblBusy.Name = "tblBusy";
            this.tblBusy.RowCount = 2;
            this.tblBusy.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 65.48672F));
            this.tblBusy.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 34.51328F));
            this.tblBusy.Size = new System.Drawing.Size(97, 15);
            this.tblBusy.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 12);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(91, 1);
            this.panel2.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Loading...";
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.White;
            this.pictureBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.pictureBox2.Image = global::ExcelPptTest.Properties.Resources.Busy;
            this.pictureBox2.Location = new System.Drawing.Point(3, 3);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(91, 3);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox2.TabIndex = 0;
            this.pictureBox2.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(949, 454);
            this.Controls.Add(this.tblSlideContainer);
            this.Name = "Form1";
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tblSlideContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picSlides)).EndInit();
            this.tblBusy.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.TableLayoutPanel tblSlideContainer;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox picSlides;
        private System.Windows.Forms.TableLayoutPanel tblBusy;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}

