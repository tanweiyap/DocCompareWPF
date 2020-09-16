﻿using DocConvert;
using DocCompareDLL;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Collections;

namespace DocCompareWPF.Classes
{
    class Document
    {
        public enum FileTypes
        {
            PDF, PPT, WORD, EXCEL, PIC, TXT, UNKNOWN
        };

        public string filePath;
        public string imageFolder;
        public FileTypes fileType;
        public List<int> docCompareIndices;
        public string docID;
        public bool loaded, processed;

        public Document()
        {
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
                
        public int ReadPPT()
        {
            PPTConvertClass pptConvertClass = new PPTConvertClass();
            int ret = -1;
            try
            {
                ret = pptConvertClass.ConvertPPTToImages(filePath, imageFolder);
                if (ret == 0)
                    processed = true;
            }
            catch
            {
                return ret;
            }

            return ret;
        }

        public void ClearFolder()
        {
            DirectoryInfo di = new DirectoryInfo(imageFolder);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }
        }

        public int ReloadDocument()
        {
            int ret = -1;
            ClearFolder();
            switch(fileType)
            {
                case FileTypes.PDF:
                    ret = ReadPDF();
                    break;
                case FileTypes.PPT:
                    ret = ReadPPT();
                    break;
                default:
                    return -1;
            }
            return ret;
        }

        public static int CompareDocs(string doc1ImageFolder, string doc2ImageFolder, string outputFolder, out ArrayList pageIndices, out int totalLen)
        {
            int ret = -1;
            int totalLength = 0;
            totalLen = new int();

            DirectoryInfo di = new DirectoryInfo(outputFolder);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }

            pageIndices = DocCompareClass.docCompare(ref doc1ImageFolder, ref doc2ImageFolder, ref outputFolder, ref totalLength);

            if (pageIndices != null) //? successful?
            {
                totalLen = totalLength;
                ret = 0;
            }

            return ret;
        }
    }
}
