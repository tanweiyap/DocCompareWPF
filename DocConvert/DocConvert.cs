using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

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
            string popplerPath = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
            popplerPath = Path.Join(popplerPath, "lib");
            //popplerPath = popplerPath.Substring(6, popplerPath.Length - 6);
            FileInfo[] fii = new DirectoryInfo(popplerPath).GetFiles();
            foreach (FileInfo f in fii)
            {
                if (f.Name.Contains("pdftoppm"))
                    popplerPath = f.FullName;
            }

            // check file existence
            if (File.Exists(filePath))
            {
                Process proc = new Process();
                proc.StartInfo.FileName = popplerPath;
                proc.StartInfo.Arguments = " -jpeg \"" + filePath + "\" \"" + outputPath + "\\page\"";
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

    public class PPTConvertClass
    {
        public PPTConvertClass()
        {
        }

        public int ConvertPPTToImages(string filePath, string outputPath, out List<bool> isHidden)
        {
            int ret = -1;
            isHidden = new List<bool>();

            try
            {
                // check file existence
                if (File.Exists(filePath))
                {
                    Type officeType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (officeType != null)
                    {
                        Application pptApplication;
                        try
                        {
                            pptApplication = new Application
                            {
                                DisplayAlerts = PpAlertLevel.ppAlertsNone, //get rid of pop ups
                                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                            };
                        }
                        catch
                        {

                            return -1;
                        }

                        Presentation pptPresentation = pptApplication.Presentations
                        .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                        , MsoTriState.msoFalse);

                        if (pptPresentation.Final) //catching another detail problem: if presentation has flag 'final' it does not allow to save it, even with SaveCopyAs... by setting it here but not saving over the original file, the original state is not changed but we can save the jpgs.
                        {
                            pptPresentation.Final = false;
                        }

                        for (int i = 1; i < (pptPresentation.Slides.Count + 1); i++)
                        {
                            pptPresentation.Slides[i].Export(outputPath + "\\" + (i - 1).ToString() + ".jpg", "jpg");

                            if (pptPresentation.Slides[i].SlideShowTransition.Hidden == MsoTriState.msoTrue)
                            {
                                isHidden.Add(true);
                            }
                            else
                            {
                                isHidden.Add(false);
                            }

                        }

                        object fileAttribute = pptPresentation.BuiltInDocumentProperties;

                        pptPresentation.Close();
                        if (pptApplication.Visible != MsoTriState.msoTrue)
                            pptApplication.Quit();
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
                    else
                    {
                        return -2;
                    }
                }

                return ret;
            }
            catch
            {
                return -1; // propably no office installation
            }
        }

        public List<string> GetFileAttribute(string filePath)
        {
            List<string> ret = new List<string>();
            try
            {
                if (File.Exists(filePath))
                {
                    Application pptApplication = new Application
                    {
                        DisplayAlerts = PpAlertLevel.ppAlertsNone, //get rid of pop ups
                        AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                    };
                    Presentation pptPresentation = pptApplication.Presentations
                    .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                    , MsoTriState.msoFalse);

                    object docProperties = pptPresentation.BuiltInDocumentProperties;
                    // Author, LastEditor, Creation Date, Modified Date
                    ret.Add(GetDocumentProperty(docProperties, "Author").ToString());
                    ret.Add(GetDocumentProperty(docProperties, "Last Author").ToString());
                    ret.Add(GetDocumentProperty(docProperties, "Creation Date").ToString());
                    ret.Add(GetDocumentProperty(docProperties, "Last Save Time").ToString());

                    pptPresentation.Close();
                    if (pptApplication.Visible != MsoTriState.msoTrue)
                        pptApplication.Quit();
                }

                return ret;
            }
            catch
            {
                return ret;
            }
        }

        private static object GetDocumentProperty(object docProperties, string propName)
        {
            object prop = docProperties.GetType().InvokeMember(
                "Item", BindingFlags.Default | BindingFlags.GetProperty,
                null, docProperties, new object[] { propName });
            object propValue = prop.GetType().InvokeMember(
                "Value", BindingFlags.Default | BindingFlags.GetProperty,
                null, prop, new object[] { });
            return propValue;
        }
    }
}