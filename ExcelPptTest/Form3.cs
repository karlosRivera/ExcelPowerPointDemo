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
        public Form3()
        {
            InitializeComponent();

            show();
        }

        private void show()
        {
            DateTime dt22 = DateTime.Now;

            string xx = dt22.ToString("yyyyMMddhhmmss");

            DataRow dr = null;
            DataTable dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("data", typeof(decimal));
            dt.Columns.Add("CountryCode", typeof(string));

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("01/01/2017");
            dr[1] = (decimal) 30 / 100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("02/01/2017");
            dr[1] = (decimal) 09/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("03/01/2017");
            dr[1] = (decimal) 15/100 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("04/01/2017");
            dr[1] = (decimal) 22/100 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("05/01/2017");
            dr[1] = (decimal) 13/100 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("06/01/2017");
            dr[1] = (decimal) 22/100 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("07/01/2017");
            dr[1] = (decimal) 07/100 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("08/01/2017");
            dr[1] = (decimal) 11/100 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("09/01/2017");
            dr[1] = (decimal) 12/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("10/01/2017");
            dr[1] = (decimal) 17 / 100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("11/01/2017");
            dr[1] = (decimal) 19/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("12/01/2017");
            dr[1] = (decimal) 02/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("13/01/2017");
            dr[1] = (decimal) 01/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("14/01/2017");
            dr[1] = (decimal) 0;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("15/01/2017");
            dr[1] = (decimal) 0;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("16/01/2017");
            dr[1] = (decimal) 07/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("17/01/2017");
            dr[1] = (decimal) 19/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("18/01/2017");
            dr[1] = (decimal) .2;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("19/01/2017");
            dr[1] = (decimal) 23/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("20/01/2017");
            dr[1] = (decimal) 35/100;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            //Chart1.BorderSkin.SkinStyle = System.Windows.Forms.DataVisualization.Charting.BorderSkinStyle.Emboss;
            ////Chart1.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            //Chart1.BorderlineColor = System.Drawing.Color.FromArgb(26, 59, 105);
            //Chart1.BorderlineWidth = 3;
            //Chart1.BackColor = Color.NavajoWhite;
            //Chart1.ChartAreas.Add("chtArea");
            Chart1.ChartAreas[0].AxisX.Title = "NPS Dates";
            Chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -60;
            Chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Verdana", 11, System.Drawing.FontStyle.Bold);
            Chart1.ChartAreas[0].AxisX.Interval = 1;

            Chart1.ChartAreas[0].AxisY.Title = "NPS Values";
            Chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Verdana", 11, System.Drawing.FontStyle.Bold);
            Chart1.ChartAreas[0].AxisY.LabelStyle.Format = "P";

            Chart1.ChartAreas[0].BorderDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid;
            Chart1.ChartAreas[0].BorderWidth = 2;

            //Chart1.Legends.Add("UnitPrice");
            //Chart1.Series.Add("UnitPricexxx");
            //Chart1.Series[0].Palette = ChartColorPalette.Bright;
            Chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            Chart1.Series[0].Points.DataBindXY(dt.DefaultView, "Date", dt.DefaultView, "Data");

            //Chart1.Series[0].IsVisibleInLegend = true;
            Chart1.Series[0].IsValueShownAsLabel = true;
            Chart1.Series[0].ToolTip = "Data Point Y Value: #VALY{G}";

            // Setting Line Width
            Chart1.Series[0].BorderWidth = 3;
            Chart1.Series[0].Color = Color.Red;

            Chart1.ChartAreas[0].AxisX.MajorGrid.LineWidth = 0;
            //Chart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 0;

            // Setting Line Shadow
            //Chart1.Series[0].ShadowOffset = 5;

            //Legend Properties
            Chart1.Legends[0].LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Table;
            Chart1.Legends[0].TableStyle = System.Windows.Forms.DataVisualization.Charting.LegendTableStyle.Wide;
            Chart1.Legends[0].Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom;
            Chart1.Width = 500;
            Chart1.Height = 300;
            Chart1.SaveImage(@"D:\test\chart.png",ChartImageFormat.Png);
        }

    }
}
