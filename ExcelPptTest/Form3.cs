﻿using System;
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
            DataRow dr = null;
            DataTable dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("data", typeof(double));
            dt.Columns.Add("CountryCode", typeof(string));

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("01/01/2017");
            dr[1] = 30 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("02/01/2017");
            dr[1] = 09;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("03/01/2017");
            dr[1] = 15 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("04/01/2017");
            dr[1] = 22 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("05/01/2017");
            dr[1] = 13 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("06/01/2017");
            dr[1] = 22 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("07/01/2017");
            dr[1] = 07 ;
            dr[2] = "GB";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("08/01/2017");
            dr[1] = 11 ;
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

            // Setting Line Shadow
            //Chart1.Series[0].ShadowOffset = 5;

            //Legend Properties
            Chart1.Legends[0].LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Table;
            Chart1.Legends[0].TableStyle = System.Windows.Forms.DataVisualization.Charting.LegendTableStyle.Wide;
            Chart1.Legends[0].Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom;
            Chart1.Width = 488;
            Chart1.Height = 345;
        }

    }
}
