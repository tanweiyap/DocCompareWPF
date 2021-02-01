using Microsoft.Office.Core;
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

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

        public int ConvertPPTToImages(string filePath, string outputPath)
        {
            int ret = -1;

            try
            {
                // check file existence
                if (File.Exists(filePath))
                {
                    Type officeType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (officeType != null)
                    {
                        PowerPoint.Application pptApplication;
                        try
                        {
                            pptApplication = new PowerPoint.Application
                            {
                                DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone, //get rid of pop ups
                                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                            };
                        }
                        catch (Exception ex)
                        {
                            return -1;
                        }

                        PowerPoint.Presentation pptPresentation = pptApplication.Presentations
                        .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                        , MsoTriState.msoFalse);

                        if (pptPresentation.Final) //catching another detail problem: if presentation has flag 'final' it does not allow to save it, even with SaveCopyAs... by setting it here but not saving over the original file, the original state is not changed but we can save the jpgs.
                        {
                            pptPresentation.Final = false;
                        }

                        for (int i = 1; i < (pptPresentation.Slides.Count + 1); i++)
                        {
                            pptPresentation.Slides[i].Export(outputPath + "\\" + (i - 1).ToString() + ".jpg", "jpg");
                        }

                        object fileAttribute = pptPresentation.BuiltInDocumentProperties;

                        pptPresentation.Close();
                        pptApplication.Quit();

                        ret = 0;
                    }
                    else
                    {
                        return -2;
                    }
                }

                return ret;
            }
            catch (Exception ex)
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
                    Type officeType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (officeType != null)
                    {
                        PowerPoint.Application pptApplication = new PowerPoint.Application
                        {
                            DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone, //get rid of pop ups
                            AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                        };
                        PowerPoint.Presentation pptPresentation = pptApplication.Presentations
                        .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                        , MsoTriState.msoFalse);

                        object docProperties = pptPresentation.BuiltInDocumentProperties;
                        // Author, LastEditor, Creation Date, Modified Date
                        ret.Add(GetDocumentProperty(docProperties, "Author").ToString());
                        ret.Add(GetDocumentProperty(docProperties, "Last Author").ToString());
                        ret.Add(GetDocumentProperty(docProperties, "Creation Date").ToString());
                        ret.Add(GetDocumentProperty(docProperties, "Last Save Time").ToString());

                        pptPresentation.Close();
                        pptApplication.Quit();
                    }
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

    public class WORDConvertClass
    {
        public WORDConvertClass()
        {
        }

        public int ConvertDoc(string filePath, out TextDocumentClass textDocument)
        {
            textDocument = new TextDocumentClass();

            try
            {
                if (File.Exists(filePath))
                {
                    Type officeType = Type.GetTypeFromProgID("Word.Application");
                    if (officeType != null)
                    {
                        Word.Application wordApplication = new Word.Application
                        {
                            DisplayAlerts = Word.WdAlertLevel.wdAlertsNone, //get rid of pop ups
                            AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                        };

                        Word.Document wordDocument = wordApplication.Documents
                        .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                        , MsoTriState.msoFalse);

                        bool trackChanges = wordDocument.TrackRevisions;

                        wordDocument.TrackRevisions = false;

                        textDocument = new TextDocumentClass(wordDocument.Paragraphs.Count);

                        foreach (Word.Paragraph para in wordDocument.Paragraphs)
                        {
                            TextParagraph thisParagraph = new TextParagraph(para.Range.Words.Count)
                            {
                                ListFormatString = para.Range.ListFormat.ListString,
                                ListFormatLevel = para.Range.ListFormat.ListLevelNumber - 1, // 0 = no indent
                                PageNumber = 0,
                                LineSpacing = para.LineSpacing
                            };

                            if (para.Range.Text != "\r")
                            {
                                if (para.Range.Font.Size == 9999999)
                                {
                                    foreach (Word.Range word in para.Range.Words)
                                    {
                                        TextWord thisWord = new TextWord
                                        {
                                            Word = word.Text.TrimEnd('\r')
                                        };
                                        thisWord.Font.FontFamily = word.Font.Name;
                                        if(word.Font.Size != 9999999)
                                            thisWord.Font.FontSize = word.Font.Size;
                                        else
                                        {
                                            foreach(Word.Range character in word.Characters)
                                            {
                                                if(character.Font.Size != 9999999)
                                                {
                                                    thisWord.Font.FontSize = character.Font.Size;
                                                    break;
                                                }
                                            }
                                        }
                                        thisWord.Font.isItalic = word.Font.Italic != 0;
                                        thisWord.Font.isBold = word.Font.Bold != 0;
                                        thisWord.Font.isUnderline = word.Font.Underline != 0;
                                        
                                        thisParagraph.Words.Add(thisWord);
                                    }
                                }
                                else
                                {
                                    TextWord thisWord = new TextWord
                                    {
                                        Word = para.Range.Text.TrimEnd('\r')
                                    };
                                    thisWord.Font.FontFamily = para.Range.Font.Name;
                                    thisWord.Font.FontSize = para.Range.Font.Size;
                                    thisWord.Font.isItalic = para.Range.Font.Italic != 0;
                                    thisWord.Font.isBold = para.Range.Font.Bold != 0;
                                    thisWord.Font.isUnderline = para.Range.Font.Underline != 0;

                                    thisParagraph.Words.Add(thisWord);
                                }
                            }
                            else
                            {
                                TextWord thisWord = new TextWord
                                {
                                    Word = ""
                                };
                                thisWord.Font.FontFamily = para.Range.Font.Name;
                                thisWord.Font.FontSize = para.Range.Font.Size;
                                thisWord.Font.isItalic = para.Range.Font.Italic != 0;
                                thisWord.Font.isBold = para.Range.Font.Bold != 0;
                                thisWord.Font.isUnderline = para.Range.Font.Underline != 0;
                                thisParagraph.Words.Add(thisWord);
                            }

                            textDocument.Paragraphs.Add(thisParagraph);
                        }

                        wordDocument.TrackRevisions = trackChanges;
                        wordDocument.Close();
                        wordApplication.Quit();
                    }

                    return 0;
                }

                return -1;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        public List<string> GetFileAttribute(string filePath)
        {
            List<string> ret = new List<string>();
            try
            {
                if (File.Exists(filePath))
                {
                    Type officeType = Type.GetTypeFromProgID("Word.Application");
                    if (officeType != null)
                    {
                        Word.Application wordApplication = new Word.Application
                        {
                            DisplayAlerts = Word.WdAlertLevel.wdAlertsNone, //get rid of pop ups
                            AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable //get rid of even more pop ups
                        };
                        Word.Document wordDocument = wordApplication.Documents
                        .Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse
                        , MsoTriState.msoFalse);

                        object docProperties = wordDocument.BuiltInDocumentProperties;
                        // Author, LastEditor, Creation Date, Modified Date
                        ret.Add(GetDocumentProperty(docProperties, "Author").ToString());
                        ret.Add(GetDocumentProperty(docProperties, "Last Author").ToString());
                        ret.Add(GetDocumentProperty(docProperties, "Creation Date").ToString());
                        ret.Add(GetDocumentProperty(docProperties, "Last Save Time").ToString());

                        wordDocument.Close();
                        wordApplication.Quit();
                    }
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

    public class TextDocumentClass
    {
        public List<TextParagraph> Paragraphs;

        public TextDocumentClass()
        {
            Paragraphs = new List<TextParagraph>();
        }

        public TextDocumentClass(int count)
        {
            Paragraphs = new List<TextParagraph>(count);
        }

        public override string ToString()
        {
            string ret = "";
            foreach (TextParagraph para in Paragraphs)
            {
                ret += para.ToString();
            }

            return ret;
        }
    }

    public class TextParagraph
    {
        public List<TextWord> Words;
        public string ListFormatString = "";
        public int ListFormatLevel;
        public double LineSpacing;
        public int PageNumber;

        public TextParagraph()
        {
            Words = new List<TextWord>();
        }

        public TextParagraph(int count)
        {
            Words = new List<TextWord>(count);
        }

        public override string ToString()
        {
            string ret = ListFormatString;

            for (uint i = 0; i < ListFormatLevel; i++)
            {
                ret += "\t";
            }

            foreach (TextWord word in Words)
            {
                ret += word.Word;
            }

            return ret;
        }
    }

    /*
    public class TextSentence
    {
        public List<TextWord> Sentence;

        public TextSentence()
        {
            Sentence = new List<TextWord>();
        }

        public TextSentence(int count)
        {
            Sentence = new List<TextWord>(count);
        }

        public override string ToString()
        {
            string ret = "";
            foreach (TextWord word in Sentence)
            {
                ret += word.Word;
            }
            return ret;
        }
    }
    */

    public class TextWord
    {
        public string Word;
        public TextFont Font;

        public TextWord()
        {
            Word = "";
            Font = new TextFont();
        }
    }

    public class TextFont
    {
        public string FontFamily = "Georgia";
        public float FontSize = 11;
        public int FontColorHex = 0;
        public bool isItalic;
        public bool isBold;
        public bool isUnderline;
    }
}