using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using OpenCvSharp;

namespace DocConvert
{
    public class PDFConvertClass
    {
        public PDFConvertClass()
        {

        }

        public int ConvertPDFtoImages(string filePath, string outputPath)
        {
            int ret = -1;
            string popplerPath = Directory.GetCurrentDirectory();
            // check file existence
            if (File.Exists(filePath))
            {
                Process proc = new Process();
                proc.StartInfo.FileName = "CMD.exe";
                proc.StartInfo.Arguments = "/c " + popplerPath + "\\pdftoppm.exe -jpeg \"" + filePath + "\" \"" + outputPath + "\\page\"";
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                proc.WaitForExit();

                DirectoryInfo di = new DirectoryInfo(outputPath);
                FileInfo[] fi = di.GetFiles();
                for (int i = 0; i < fi.Length; i++)
                {
                    string[] filename = fi[i].Name.Split("page-");
                    string name = Path.GetFileNameWithoutExtension(filename[1]);
                    int pageName = int.Parse(name);
                    File.Move(fi[i].FullName, Path.Join(outputPath, (pageName - 1).ToString() + ".jpg"));
                }
                ret = 0;
            }

            return ret;
        }

        public int ConvertPPTToImages(string filePath, string outputPath)
        {
            int ret = -1;

            // check file existence
            if (File.Exists(filePath))
            {
                Application pptApplication = new Application();
                Presentation pptPresentation = pptApplication.Presentations
                .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                , MsoTriState.msoFalse);

                pptPresentation.Export(outputPath, "jpg", Int32.Parse(pptPresentation.SlideMaster.Width.ToString()), Int32.Parse(pptPresentation.SlideMaster.Height.ToString()));
                /*
                DirectoryInfo di = new DirectoryInfo(outputPath);
                FileInfo[] fi = di.GetFiles();
                for (int i = 0; i < fi.Length; i++)
                {
                    string[] filename = fi[i].Name.Split("page-");
                    string name = Path.GetFileNameWithoutExtension(filename[1]);
                    int pageName = int.Parse(name);
                    File.Move(fi[i].FullName, Path.Join(outputPath, (pageName - 1).ToString() + ".jpg"));
                }
                */
                ret = 0;
            }

            return ret;
        }

    }

    public class PPTConvertClass
    {
        public PPTConvertClass()
        {

        }

        public int ConvertPPTToImages(string filePath, string outputPath)
        {
            int ret = -1;

            // check file existence
            if (File.Exists(filePath))
            {
                Application pptApplication = new Application
                {
                    DisplayAlerts = Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone, //get rid of pop ups
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                };
                Presentation pptPresentation = pptApplication.Presentations
                .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                , MsoTriState.msoFalse);

                if (pptPresentation.Final) //catching another detail problem: if presentation has flag 'final' it does not allow to save it, even with SaveCopyAs... by setting it here but not saving over the original file, the original state is not changed but we can save the jpgs.
                {
                    pptPresentation.Final = false;
                }
                
                for (int i = 1; i < (pptPresentation.Slides.Count + 1); i++)
                {
                    pptPresentation.Slides[i].Export(outputPath + "\\" + (i-1).ToString() + ".jpg", "jpg");
                }
                pptPresentation.Close();

                //old code by WYT...
                //pptPresentation.Export(outputPath, "jpg", Int32.Parse(pptPresentation.SlideMaster.Width.ToString()), Int32.Parse(pptPresentation.SlideMaster.Height.ToString()));
                
                //DirectoryInfo di = new DirectoryInfo(outputPath);
                //FileInfo[] fi = di.GetFiles();
                //for (int i = 0; i < fi.Length; i++)
                //{
                //    string[] filename;
                //    if (fi[i].Name.Contains("Folie"))
                //        filename = fi[i].Name.Split("Folie");
                //    else
                //        filename = fi[i].Name.Split("Slide");

                //    string name = Path.GetFileNameWithoutExtension(filename[1]);
                //    int pageName = int.Parse(name);
                //    File.Move(fi[i].FullName, Path.Join(outputPath, (pageName - 1).ToString() + ".jpg"));
                //}
                
                ret = 0;
            }

            return ret;
        }

    }

    public class PICConvertClass
    {
        public PICConvertClass()
        {

        }

        public int CovertPICtoJPEG(string filePath, string outputPath)
        {
            int ret = -1;

            try
            {
                Mat thisImage = Cv2.ImRead(filePath);
                thisImage.SaveImage(Path.Join(outputPath, "0.jpg"));
            }
            catch
            {
                return ret;
            }

            ret = 0;

            return ret;
        }
    }
}
