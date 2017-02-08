using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;

namespace ExcelPptTest
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string fileName = @"C:\Users\Tridip\Desktop\KPIDashBoard20170125 125513.pptx";
            string exportName = "video_of_presentation";
            string exportPath = @"D:\test\Export\{0}.wmv";

            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            ppApp.WindowState = PpWindowState.ppWindowMinimized;
            Microsoft.Office.Interop.PowerPoint.Presentations oPresSet = ppApp.Presentations;
            Microsoft.Office.Interop.PowerPoint._Presentation oPres = oPresSet.Open(fileName,
                        Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoFalse);
            try
            {
                //oPres.CreateVideo(exportName);
                oPres.SaveCopyAs(String.Format(exportPath, exportName),
                    Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsEMF,
                    Microsoft.Office.Core.MsoTriState.msoCTrue);
            }
            finally
            {
                ppApp.Quit();
            }
        }
    }
}
