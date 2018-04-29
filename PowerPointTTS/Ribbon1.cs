using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.IO;

namespace PowerPointTTS
{
    public partial class Ribbon1
    {
        private MainFormTTS mainFomr = null;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (mainFomr != null)
            {
                mainFomr.Close();
                mainFomr = null;
                
            }

            mainFomr = new MainFormTTS();
            mainFomr.Show();

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string appPath = Directory.GetCurrentDirectory();
            string strtemp = Globals.ThisAddIn.Application.ActivePresentation.Name.Replace(".pptx", "");
            appPath = appPath + @"\" + strtemp + @"\Recode\";

            try
            {
                System.Diagnostics.Process.Start(appPath);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("録音ファイルまだございません。");
            }
        }
    }
}
