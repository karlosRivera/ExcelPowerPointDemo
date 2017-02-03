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
    public partial class Form2 : Form
    {
        public Form2()
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
            decimal decval = 0;

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

        private void Form2_Load(object sender, EventArgs e)
        {
            pptNS.ApplicationClass powerpointApplication = null;
            pptNS.Presentation pptPresentation = null;
            pptNS.Slide pptSlide = null;
            pptNS.ShapeRange shapeRange = null;
            pptNS.Shape oTxtShape = null;
            pptNS.Shape oShape = null;
            pptNS.Shape oChartShape = null;


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
                pptSlide = pptPresentation.Slides.Add(1, pptNS.PpSlideLayout.ppLayoutBlank);


                // capture range
                //var writeRange = targetSheet.Range["A1:B15"];
                destRange = targetSheet.get_Range("A1:B15");

                // adding header text for slides
                oTxtShape = pptSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left: 30, Top: 30, Width: 340, Height: 340);
                oTxtShape.TextFrame.TextRange.Text = "This is my demo text";
                oTxtShape.TextEffect.FontName = "Arial";
                oTxtShape.TextEffect.FontSize = 32;
                oTxtShape.TextEffect.Alignment = Microsoft.Office.Core.MsoTextEffectAlignment.msoTextEffectAlignmentCentered;

                System.Array myvalues = (System.Array)destRange.Cells.Value;
                List<Tuple<string, string>> cellData = GetData(myvalues);

                int iRows = cellData.Count + 1;
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


                    oShape.Table.Cell(row, 2).Shape.TextFrame.TextRange.Text = (strValue.StartsWith("0") ? "0%" : (strValue + "0%"));
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

                oShape.Top = 100;
                oShape.Left = 30;

               oChartShape= pptSlide.Shapes.AddChart(Microsoft.Office.Core.XlChartType.xlLine, 20F, 30F, 400F, 300F);
               

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
    }
}
