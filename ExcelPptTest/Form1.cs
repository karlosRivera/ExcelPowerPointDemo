using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using xlNS = Microsoft.Office.Interop.Excel;
using pptNS = Microsoft.Office.Interop.PowerPoint;

namespace ExcelPptTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<Tuple<string, string>> GetData(System.Array values)
        {

            // create a new string array
            //string[] theArray = new string[values.Length];

            //// loop through the 2-D System.Array and populate the 1-D String Array
            //for (int i = 1; i <= values.Length; i++)
            //{
            //    if (values.GetValue(1, i) == null)
            //        theArray[i - 1] = "";
            //    else
            //        theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            //}

            var data = new List<Tuple<string, string>>();
            string strDate = "", strValue = "";
            decimal decval=0;

            for (int i = 2; i <= 15; i++)
            {
                for (int j = 1; j <= 2; j++)
                {
                    if (j > 1)
                    {
                        if (values.GetValue(i, j - 1) == null)
                        {
                            strDate = "";
                        }
                        else
                        {
                            strDate = DateTime.Parse(values.GetValue(i, j - 1).ToString()).ToString("dd/MM/yyyy");
                        }

                        if (values.GetValue(i, j) == null)
                        {
                            strValue = "";
                        }
                        else
                        {
                            decval = decimal.Parse(values.GetValue(i, j).ToString()) * 10;
                            strValue = Convert.ToInt32(decval).ToString();
                        }

                        data.Add(new Tuple<string, string>(strDate, strValue));
                    }
                }
            }

            return data;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pptNS.ApplicationClass powerpointApplication = null;
            pptNS.Presentation pptPresentation = null;
            pptNS.Slide pptSlide = null;
            pptNS.ShapeRange shapeRange = null;
            pptNS.Shape oShape = null;

            xlNS.ApplicationClass excelApplication = null;
            xlNS.Workbook excelWorkBook = null;
            xlNS.Worksheet targetSheet = null;
            xlNS.ChartObjects chartObjects = null;
            xlNS.ChartObject existingChartObject = null;
            xlNS.Range destRange = null;

            string paramPresentationPath = @"D:\test\Chart Slide.pptx";
            string paramWorkbookPath = @"D:\test\NPS.xlsx";
            object paramMissing = Type.Missing;


            try
            {
                // Create an instance of PowerPoint.
                powerpointApplication = new pptNS.ApplicationClass();

                // Create an instance Excel.          
                excelApplication = new xlNS.ApplicationClass();

                // Open the Excel workbook containing the worksheet with the chart
                // data.
                excelWorkBook = excelApplication.Workbooks.Open(paramWorkbookPath,
                                paramMissing, paramMissing, paramMissing,
                                paramMissing, paramMissing, paramMissing,
                                paramMissing, paramMissing, paramMissing,
                                paramMissing, paramMissing, paramMissing,
                                paramMissing, paramMissing);

                // Get the worksheet that contains the chart.
                targetSheet =
                    (xlNS.Worksheet)(excelWorkBook.Worksheets["Spain"]);

                // Get the ChartObjects collection for the sheet.
                chartObjects =
                    (xlNS.ChartObjects)(targetSheet.ChartObjects(paramMissing));



                // Create a PowerPoint presentation.
                pptPresentation = powerpointApplication.Presentations.Add(
                                    Microsoft.Office.Core.MsoTriState.msoTrue);

                // Add a blank slide to the presentation.
                pptSlide =
                    pptPresentation.Slides.Add(1, pptNS.PpSlideLayout.ppLayoutBlank);

                // capture range
                //var writeRange = targetSheet.Range["A1:B15"];
                destRange = targetSheet.get_Range("A1:B15");

                pptSlide.Shapes[1].TextFrame.TextRange.Text = "Spain Data 01/01/2017 to 20/01/2017";
                pptSlide.Shapes[1].TextFrame.TextRange.Font.Name = "Verdana";
                pptSlide.Shapes[1].TextFrame.TextRange.Font.Size = 10;

                System.Array myvalues = (System.Array)destRange.Cells.Value;
                List<Tuple<string, string>> cellData = GetData(myvalues);

                int iRows = cellData.Count+1;
                int iColumns = 2;
                int row = 2;

                oShape = pptSlide.Shapes.AddTable(iRows, iColumns, 500, 110, 160, 120);
                oShape.Table.Cell(1, 1).Merge(oShape.Table.Cell(1, 2));

                oShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Spain Data 01/01/2017 to 20/01/2017";
                oShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Name = "Verdana";
                oShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size = 8;

                foreach (Tuple<string, string> item in cellData)
                {
                    string strdate = item.Item1;
                    string strValue = item.Item2;

                    oShape.Table.Cell(row, 1).Shape.TextFrame.TextRange.Text = strdate;
                    oShape.Table.Cell(row, 1).Shape.TextFrame.TextRange.Font.Name = "Verdana";
                    oShape.Table.Cell(row, 1).Shape.TextFrame.TextRange.Font.Size = 8;


                    oShape.Table.Cell(row, 2).Shape.TextFrame.TextRange.Text = (strValue.StartsWith("0") ?  "0%" : (strValue + "0%"));
                    oShape.Table.Cell(row, 2).Shape.TextFrame.TextRange.Font.Name = "Verdana";
                    oShape.Table.Cell(row, 2).Shape.TextFrame.TextRange.Font.Size = 8;

                    //if (row == 1)
                    //{
                    //    oShape.Table.Cell(row, 1).Shape.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(208, 208, 208).ToArgb();
                    //    oShape.Table.Cell(row, 1).Shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                    //    oShape.Table.Cell(row, 2).Shape.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(208, 208, 208).ToArgb();
                    //    oShape.Table.Cell(row, 2).Shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                    //}

                    row++;
                }

                oShape.Top = 10;
                oShape.Left =10;

                //copy range
                //destRange.Copy();

                // Paste the chart into the PowerPoint presentation.
                //shapeRange = pptSlide.Shapes.Paste();

                //var table = pptSlide.Shapes.AddTable();
                // Position the chart on the slide.
                //shapeRange.Left = 60;
                //shapeRange.Top = 100;

                // Get or capture the chart to copy.
                //existingChartObject = (xlNS.ChartObject)(chartObjects.Item(1));


                // Copy the chart from the Excel worksheet to the clipboard.
                //existingChartObject.Copy();

                // Paste the chart into the PowerPoint presentation.
                //shapeRange = pptSlide.Shapes.Paste();
                //Position the chart on the slide.
                //shapeRange.Left = 90;
                //shapeRange.Top = 100;

                // Save the presentation.
                pptPresentation.SaveAs(paramPresentationPath,
                                pptNS.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                                Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Release the PowerPoint slide object.
                shapeRange = null;
                pptSlide = null;

                // Close and release the Presentation object.
                if (pptPresentation != null)
                {
                    pptPresentation.Close();
                    pptPresentation = null;
                }

                // Quit PowerPoint and release the ApplicationClass object.
                if (powerpointApplication != null)
                {
                    powerpointApplication.Quit();
                    powerpointApplication = null;
                }

                // Release the Excel objects.
                targetSheet = null;
                chartObjects = null;
                existingChartObject = null;

                // Close and release the Excel Workbook object.
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }

                // Quit Excel and release the ApplicationClass object.
                if (excelApplication != null)
                {
                    excelApplication.Quit();
                    excelApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Data", typeof(Int32));

            DataRow dr = dt.NewRow();
            dr[0] = DateTime.Parse("01-03-2017");
            dr[1] = 20;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("01-04-2017");
            dr[1] = 0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("01-05-2017");
            dr[1] = 0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = DateTime.Parse("01-09-2017");
            dr[1] = 10;
            dt.Rows.Add(dr);

            DateTimeOffset localTime = DateTimeOffset.UtcNow;
            var offsetVal = localTime.Offset;

        }

        //        private void button1_Click(object sender, EventArgs e)
        //        {
        //            pptNS.Presentation _pPres = null;
        //            pptNS.ApplicationClass _pApp = null;
        //            int noofSalespersons = 0 ;

        //            _pApp = new pptNS.ApplicationClass();

        ////Graph.Chart objChart;

        //int noofSlides;

        //noofSalespersons =12;

        //string sTemplateFile = Server.MapPath("Resource\\Template.ppt");

        //_pPres = _pApp.Presentations.Open(sTemplateFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

        //noofSlides = (noofSalespersons / 10) + 1;

        //if (noofSalespersons%10>0)

        //{

        //noofSlides += 1;

        //}

        //int slideIndex=1;

        //while (_pPres.Slides.Count!=noofSlides)

        //{

        //_pPres.Slides.InsertFromFile("D:\\Template.ppt", slideIndex, 1, 1);

        //slideIndex += 1;

        //}


        //Chart chart;

        //Graph.DataSheet dataSheet;

        //for (int i = 1; i <= slideIndex; i++)

        //{

        //pptNS.Slide slide = _pPres.Slides._Index(i) as Slide;

        //if (i==1)

        //{

        //pptNS.Shape shape = slide.Shapes[1];


        //chart = shape.OLEFormat.Object as Graph.Chart;

        //dataSheet = chart.Application.DataSheet;

        //dataSheet.Cells[2, 2] = "50";

        //dataSheet.Cells[2, 3] = "40";

        //dataSheet.Cells[2, 4] = "50";

        //dataSheet.Cells[2, 5] = "50";

        //dataSheet.Cells[3, 2] = "60";

        //dataSheet.Cells[3, 3] = "70";

        //dataSheet.Cells[3, 4] = "80";

        //dataSheet.Cells[3, 5] = "60";

        ////dataSheet.Cells[3, 6] = "0";

        //dataSheet.Cells[4, 2] = "50";

        //dataSheet.Cells[4, 3] = "40";

        //dataSheet.Cells[4, 4] = "50";

        //dataSheet.Cells[4, 5] = "50";

        ////dataSheet.Cells[4, 6] = "0";

        //chart.Application.Update();

        //dataSheet = null;

        //chart = null;

        //}

        //ArrayList listShapeName = new ArrayList();

        //ArrayList listAutoShapes = new ArrayList();

        //ArrayList listTextBox = new ArrayList();

        //ArrayList listPlaceHolder = new ArrayList();

        //int reminingShapes;

        //reminingShapes = noofSalespersons % 10;

        //if (i>=2)

        //{

        //if ((i == slideIndex) && (reminingShapes>0))

        //{

        //int remingCount = 0;

        //for (int k = 1; k <= slide.Shapes.Count; k++)

        //{

        //pptNS.Shape shape = slide.Shapes[k];

        ////string shapeName = shape.Name;

        ////listShapeName.Add(shapeName);

        //if (shape.Type == MsoShapeType.msoEmbeddedOLEObject)

        //{

        //remingCount += 1;

        //if (remingCount>reminingShapes)

        //{

        ////slide.Shapes[k].Delete();

        //listShapeName.Add(k);

        //}

        //}

        //}

        //for (int m = listShapeName.Count - 1; m >= 0; m--)

        //{

        //slide.Shapes[listShapeName[m]].Delete();

        //}

        //listShapeName.Clear();

        //for (int k = 1; k <= slide.Shapes.Count; k++)

        //{

        //pptNS.Shape shape = slide.Shapes[k];


        //if (reminingShapes < 6)

        //{

        //if (shape.Type == MsoShapeType.msoTextBox)

        //{

        //listTextBox.Add(k);

        //}

        //}

        //}

        //if (listTextBox.Count>0)

        //{

        //for (int m = listTextBox.Count - 1; m >= listTextBox.Count - 4; m--)

        //{

        //slide.Shapes[listTextBox[m]].Delete();

        //}

        //listTextBox.Clear();

        //}


        //for (int k = 1; k <= slide.Shapes.Count; k++)

        //{

        //pptNS.Shape shape = slide.Shapes[k];

        //if (reminingShapes < 6)

        //{

        //if (shape.Type == MsoShapeType.msoAutoShape)

        //{

        //listAutoShapes.Add(k);

        //}

        //}

        //}

        //if (listAutoShapes.Count>0)

        //{

        //for (int m = listAutoShapes.Count - 1; m >= listAutoShapes.Count - 4; m--)

        //{

        //slide.Shapes[listAutoShapes[m]].Delete();

        //}

        //listAutoShapes.Clear();

        //}

        //}

        //int z = 0;

        //int bottom = 0;

        //for (int k = 1; k <= slide.Shapes.Count; k++)

        //{

        //pptNS.Shape shape = slide.Shapes[k];

        //string shapeName= shape.Name;

        //if (shape.Type==MsoShapeType.msoEmbeddedOLEObject)

        //{

        //listShapeName.Add(shapeName);


        //chart =shape.OLEFormat.Object as Graph.Chart;

        //dataSheet = chart.Application.DataSheet;

        ////if (k % 2 == 0)

        ////{

        //z = z + 1;


        ////shape.Top = newHeight;

        ////shape.IncrementTop(200);//(200-shape.Height) + shape.Top);

        //shape.Height = ((42 - (2 * z)) * 200) / (40);

        ////shape.Width = 90;

        //dataSheet.Cells[2, 2] = (((11-z) * 100) / 100).ToString();

        //dataSheet.Cells[3, 2] = (((11-z) * 100) / 100).ToString();

        //dataSheet.Cells[4, 2] = (((11-z) * 100) / 100).ToString();

        //dataSheet.Cells[5, 2] = (((11-z) * 100) / 100).ToString();

        ////shape.Width = 90;


        ////}

        ////else

        ////{


        //// //shape.Width = 86;

        //// shape.Height = 200;

        //// //newHeight = shape.Top;

        //// dataSheet.Cells[2, 2] = ((80 * 100) / 260).ToString();

        //// dataSheet.Cells[3, 2] = ((70 * 100) / 260).ToString();

        //// dataSheet.Cells[4, 2] = ((60 * 100) / 260).ToString();

        //// dataSheet.Cells[5, 2] = ((50 * 100) / 260).ToString();

        ////}

        //chart.Application.Update();

        //dataSheet = null;

        //chart = null;

        //}

        //}




        //}

        //_pPres.SaveAs("d:\\Vijay\\pptNS\\Sample2.ppt", PpSaveAsFileType.ppSaveAsPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);

        //_pPres.Close();

        //_pApp.Quit();

        //GC.Collect();

        //        }
        //    }
        //}

    }

}
