using DocCompareDLL;
using DocConvert;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace DocCompareWPF.Classes
{
    internal class Document
    {
        public string CreatedDate;
        public string Creator;
        public List<int> docCompareIndices;
        public string docID;
        public string filePath;
        public FileTypes fileType;
        public string imageFolder;
        public string LastEditor;
        public bool loaded, processed;
        public string ModifiedDate;
        public List<bool> pptIsHidden;

        public Document()
        {
        }

        public enum FileTypes
        {
            PDF, PPT, WORD, EXCEL, PIC, TXT, UNKNOWN
        };

        public static int CompareDocs(string doc1ImageFolder, string doc2ImageFolder, string outputFolder, out ArrayList pageIndices, out int totalLen, int[,] forceIndices)
        {
            int ret = -1;
            int totalLength = 0;
            totalLen = new int();

            DirectoryInfo di = new DirectoryInfo(outputFolder);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }

            pageIndices = DocCompareClass.DocCompare(ref doc1ImageFolder, ref doc2ImageFolder, ref outputFolder, ref totalLength, forceIndices);

            if (pageIndices != null) //? successful?
            {
                totalLen = totalLength;
                ret = 0;
            }

            return ret;
        }

        public void DetectFileType()
        {
            string extension = Path.GetExtension(filePath);
            fileType = extension switch
            {
                ".ppt" => FileTypes.PPT,
                ".pptx" => FileTypes.PPT,
                ".PPT" => FileTypes.PPT,
                ".PPTX" => FileTypes.PPT,
                ".doc" => FileTypes.WORD,
                ".docx" => FileTypes.WORD,
                ".DOC" => FileTypes.WORD,
                ".DOCX" => FileTypes.WORD,
                ".xls" => FileTypes.EXCEL,
                ".xlsx" => FileTypes.EXCEL,
                ".XLS" => FileTypes.EXCEL,
                ".XLSX" => FileTypes.EXCEL,
                ".pdf" => FileTypes.PDF,
                ".PDF" => FileTypes.PDF,
                ".txt" => FileTypes.TXT,
                ".jpg" => FileTypes.PIC,
                ".jpeg" => FileTypes.PIC,
                ".JPG" => FileTypes.PIC,
                ".JPEG" => FileTypes.PIC,
                ".bmp" => FileTypes.PIC,
                ".BMP" => FileTypes.PIC,
                ".png" => FileTypes.PIC,
                ".PNG" => FileTypes.PIC,
                ".gif" => FileTypes.PIC,
                ".GIF" => FileTypes.PIC,
                _ => FileTypes.UNKNOWN,
            };
        }

        public int ReadPDF()
        {
            PDFConvertClass pdfClass = new PDFConvertClass();
            int ret = pdfClass.ConvertPDFtoImages(filePath, imageFolder);
            if (ret == 0)
                processed = true;
            return ret;
        }

        public int ReadPic()
        {
            PICConvertClass picConvertClass = new PICConvertClass();
            int ret = -1;
            try
            {
                ret = picConvertClass.CovertPICtoPNG(filePath, imageFolder);
                if (ret == 0)
                    processed = true;
            }
            catch
            {
                return ret;
            }
            return ret;
        }

        public int ReadPPT()
        {
            PPTConvertClass pptConvertClass = new PPTConvertClass();
            int ret = -1;
            try
            {
                ret = pptConvertClass.ConvertPPTToImages(filePath, imageFolder, out pptIsHidden);
                if (ret == 0)
                    processed = true;
            }
            catch
            {
                return ret;
            }

            return ret;
        }

        public void ReadStats(string cultureInfo) // must be called after DetectFileType()
        {
            CultureInfo culture = CultureInfo.GetCultureInfo(cultureInfo);
            FileInfo fileInfo;
            switch (fileType)
            {
                case FileTypes.PPT:
                    List<string> fileAttributes = new PPTConvertClass().GetFileAttribute(filePath);
                    if (fileAttributes.Count == 4) // read successfully
                    {
                        Creator = fileAttributes[0];
                        LastEditor = fileAttributes[1];
                        CreatedDate = DateTime.Parse(fileAttributes[2]).ToString("F", culture);
                        ModifiedDate = DateTime.Parse(fileAttributes[3]).ToString("F", culture);
                    }

                    if (Creator == null || LastEditor == null || CreatedDate == null || ModifiedDate == null)
                    {
                        fileInfo = new FileInfo(filePath);
                        Creator = fileInfo.GetAccessControl().GetOwner(typeof(System.Security.Principal.NTAccount)).ToString().Split("\\")[^1];
                        LastEditor = Creator;
                        CreatedDate = fileInfo.CreationTime.ToString("F", culture);
                        ModifiedDate = fileInfo.LastWriteTime.ToString("F", culture);
                    }
                    break;

                default:
                    fileInfo = new FileInfo(filePath);
                    Creator = fileInfo.GetAccessControl().GetOwner(typeof(System.Security.Principal.NTAccount)).ToString().Split("\\")[^1];
                    LastEditor = Creator;
                    CreatedDate = fileInfo.CreationTime.ToString("F", culture);
                    ModifiedDate = fileInfo.LastWriteTime.ToString("F", culture);
                    break;
            }
        }
        public int ReloadDocument(string workingDir)
        {
            int ret;

            docID = Guid.NewGuid().ToString();
            imageFolder = Path.Join(workingDir, docID);
            Directory.CreateDirectory(imageFolder);

            switch (fileType)
            {
                case FileTypes.PDF:
                    ret = ReadPDF();
                    break;

                case FileTypes.PPT:
                    ret = ReadPPT();
                    break;

                case FileTypes.PIC:
                    ret = ReadPic();
                    break;

                default:
                    return -1;
            }
            return ret;
        }
    }
}