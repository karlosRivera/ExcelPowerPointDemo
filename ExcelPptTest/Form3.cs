using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;


namespace ExcelPptTest
{
    public partial class Form3 : Form
    {
        int _x = -1;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private RectangleF _rect;


        public Form3()
        {
            InitializeComponent();
            Init();
        }

        private void Init()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint1 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 1D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint2 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(1D, 2D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint3 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(2D, 4D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint4 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(3D, 5D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint5 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(4D, 4D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint6 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(5D, 1D);
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();

            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(32, 45);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);
            series1.Points.Add(dataPoint6);
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(584, 360);
            this.chart1.TabIndex = 0;
            this.chart1.Text = "chart1";
            this.chart1.PostPaint += new System.EventHandler<System.Windows.Forms.DataVisualization.Charting.ChartPaintEventArgs>(this.chart1_PostPaint);
            this.chart1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chart1_KeyDown);

            this.Controls.Add(this.chart1);
            this.ClientSize = new System.Drawing.Size(676, 474);

            this.chart1.MouseDown += new MouseEventHandler(chart1_MouseDown);

            this.Text = "Use Ctrl and Arrow keys to move the rectangle";
        }

        void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && this._rect.Contains(new Point(e.X, e.Y)))
                MessageBox.Show("inside");
        }

        //controlKey must be held down
        private void chart1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Right | Keys.Control))
            {
                _x++;

                _x = Math.Min(this.chart1.Series[0].Points.Count, _x);

                this.chart1.Invalidate();
            }

            if (e.KeyData == (Keys.Left | Keys.Control))
            {
                _x--;

                _x = Math.Max(_x, -1);

                this.chart1.Invalidate();
            }
        }

        private void chart1_PostPaint(object sender, System.Windows.Forms.DataVisualization.Charting.ChartPaintEventArgs e)
        {
            object o = e.ChartElement;

            if (o.GetType().Equals(typeof(System.Windows.Forms.DataVisualization.Charting.Series)))
            {
                Series s = (Series)o;

                //maxima and minima
                double y2 = chart1.ChartAreas[0].AxisY.Maximum;
                double y1 = chart1.ChartAreas[0].AxisY.Minimum;

                float xVal = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.X, _x);
                float xVal2 = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.X, _x - 1);
                float yVal = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.Y, y1);
                float yVal2 = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.Y, y2);

                //rectangle for the moveable "Pointer"
                this._rect = e.ChartGraphics.GetAbsoluteRectangle(new RectangleF(new PointF(xVal - 0.5F, yVal2), new SizeF(1F, Math.Abs(yVal - yVal2))));

                //draw the pointer
                using (SolidBrush sb = new SolidBrush(Color.FromArgb(64, 255, 0, 0)))
                    e.ChartGraphics.Graphics.FillRectangle(sb, Rectangle.Round(this._rect));

                //check the points of the current series
                for (int i = 0; i < s.Points.Count; i++)
                {
                    //some values needed
                    float xV = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.X, s.Points[i].XValue);
                    float xV2 = 0;
                    if (i > 0)
                        xV2 = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.X, s.Points[i - 1].XValue);
                    float yV = (float)e.ChartGraphics.GetPositionFromAxis(e.Chart.ChartAreas[0].Name, System.Windows.Forms.DataVisualization.Charting.AxisName.Y, s.Points[i].YValues[0]);

                    //if the yValue is greate than 3 draw a red rectangle
                    if (s.Points[i].YValues[0] > 3)
                    {
                        RectangleF r = e.ChartGraphics.GetAbsoluteRectangle(
                            new RectangleF(new PointF(xV - Math.Abs(xV - xV2) / 2F + 1, yV),
                                new SizeF(Math.Abs((xV - 1) - (xV2 + 1)), yVal - yV)));
                        e.ChartGraphics.Graphics.FillRectangle(Brushes.Red, r);
                    }
                }
            }
        }
    }
}
