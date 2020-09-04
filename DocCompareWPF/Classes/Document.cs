using DocConvert;
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

        public Document()
        {
        }

        public void detectFileType()
        {
            string extension = Path.GetExtension(filePath);
            switch(extension)
            {
                case ".ppt":
                    fileType = FileTypes.PPT;
                    break;
                case ".pptx":
                    fileType = FileTypes.PPT;
                    break;
                case ".PPT":
                    fileType = FileTypes.PPT;
                    break;
                case ".PPTX":
                    fileType = FileTypes.PPT;
                    break;
                case ".doc":
                    fileType = FileTypes.WORD;
                    break;
                case ".docx":
                    fileType = FileTypes.WORD;
                    break;
                case ".DOC":
                    fileType = FileTypes.WORD;
                    break;
                case ".DOCX":
                    fileType = FileTypes.WORD;
                    break;
                case ".xls":
                    fileType = FileTypes.EXCEL;
                    break;
                case ".xlsx":
                    fileType = FileTypes.EXCEL;
                    break;
                case ".XLS":
                    fileType = FileTypes.EXCEL;
                    break;
                case ".XLSX":
                    fileType = FileTypes.EXCEL;
                    break;
                case ".pdf":
                    fileType = FileTypes.PDF;
                    break;
                case ".PDF":
                    fileType = FileTypes.PDF;
                    break;
                case ".txt":
                    fileType = FileTypes.TXT;
                    break;
                case ".jpg":
                    fileType = FileTypes.PIC;
                    break;
                case ".jpeg":
                    fileType = FileTypes.PIC;
                    break;
                case ".JPG":
                    fileType = FileTypes.PIC;
                    break;
                case ".JPEG":
                    fileType = FileTypes.PIC;
                    break;
                case ".bmp":
                    fileType = FileTypes.PIC;
                    break;
                case ".BMP":
                    fileType = FileTypes.PIC;
                    break;
                case ".png":
                    fileType = FileTypes.PIC;
                    break;
                case ".PNG":
                    fileType = FileTypes.PIC;
                    break;
                case ".gif":
                    fileType = FileTypes.PIC;
                    break;
                case ".GIF":
                    fileType = FileTypes.PIC;
                    break;
                default:
                    fileType = FileTypes.UNKNOWN;
                    break;
            }
        }

        public int readPDF()
        {            
            PDFConvertClass pdfClass = new PDFConvertClass();
            int ret = pdfClass.convertPDFtoImages(filePath, imageFolder);

            return ret;
        }
                
        public void clearFolder()
        {
            DirectoryInfo di = new DirectoryInfo(imageFolder);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }
        }

        public static int compareDocs(string doc1ImageFolder, string doc2ImageFolder, string outputFolder, out ArrayList pageIndices, out int totalLen)
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
