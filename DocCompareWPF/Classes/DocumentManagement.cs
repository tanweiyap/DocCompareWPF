using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace DocCompareWPF.Classes
{
    class DocumentManagement
    {
        public List<Document> documents;
        public int MAX_DOC_COUNT = 5;
        public List<int> documentsToShow;
        public List<int> documentsToCompare;
        public ArrayList pageCompareIndices;
        public int totalLen;
        public string workingDir;

        public DocumentManagement(string p_workingDir, AppSettings settings)
        {
            documents = new List<Document>();
            if(settings.numPanelsDragDrop == 3)
                documentsToShow = new List<int>() { 0, 1, 2 };
            else
                documentsToShow = new List<int>() { 0, 1};

            documentsToCompare = new List<int>() { 0, 1 };
            workingDir = p_workingDir;

            DirectoryInfo di = new DirectoryInfo(workingDir);
            di.Delete(true);
            Directory.CreateDirectory(workingDir);

            Directory.CreateDirectory(Path.Join(workingDir, "compare"));
        }

        public DocumentManagement(int p_maxDocCount, string p_workingDir, AppSettings settings)
        {
            MAX_DOC_COUNT = p_maxDocCount;
            documents = new List<Document>();
            if (settings.numPanelsDragDrop == 3)
                documentsToShow = new List<int>() { 0, 1, 2 };
            else
                documentsToShow = new List<int>() { 0, 1 };
            documentsToCompare = new List<int>() { 0, 1 };
            workingDir = p_workingDir;

            DirectoryInfo di = new DirectoryInfo(workingDir);
            di.Delete(true);
            Directory.CreateDirectory(workingDir);

            Directory.CreateDirectory(Path.Join(workingDir, "compare"));
        }

        public void AddDocument(string p_filePath)
        {
            Document document = new Document
            {
                filePath = p_filePath,
                docID = System.Guid.NewGuid().ToString()
            };
            document.imageFolder = Path.Join(workingDir, document.docID);
            Directory.CreateDirectory(document.imageFolder);
            documents.Add(document);
        }

        public void RemoveDocument(int index, int viewID)
        {
            //string[] docShown = new string[3] { documents[documentsToShow[0]].docID, documents[documentsToShow[1]].docID, documents[documentsToShow[2]].docID };
            List<string> docShown = new List<string>();
            for(int i = 0; i < documentsToShow.Count; i++)
            {
                if(documentsToShow[i] != -1)
                {
                    docShown.Add(documents[documentsToShow[i]].docID);
                }
            }

            // clean up folder
            documents[index].ClearFolder();
            DirectoryInfo di = new DirectoryInfo(documents[index].imageFolder);
            di.Delete();
            documents.RemoveAt(index);

            // we now check for docs to show
            for(int i = 0; i< docShown.Count; i++)
            {
                if(i == viewID) // view, where the doc was removed
                {
                    string docIDToConsider;
                    for(int j = 0; j < documents.Count; j++)
                    {
                        docIDToConsider = documents[j].docID;
                        bool ok = false;
                        for(int k = 0; k < docShown.Count; k++)
                        {
                            if (k != viewID)
                            {
                                if (docIDToConsider == docShown[k])
                                {
                                    ok = false;
                                    break;
                                }
                                else
                                {
                                    ok = true;
                                }
                            }
                        }

                        if (ok == true)
                        {
                            documentsToShow[i] = j;
                            break;
                        }
                        else
                        {
                            documentsToShow[i] = -1;
                        }
                    }
                }
                else
                {
                    documentsToShow[i] = documents.FindIndex(x => x.docID == docShown[i]);
                }
            }

            if(documents.Count == 0)
            {
                documentsToShow[0] = -1;
            }

            for(int i = 0; i < 3; i++)
            {
                if(documentsToShow[i] == -1)
                {
                    documentsToShow.RemoveAt(i);
                    documentsToShow.Add(-1);
                }
            }
        }
    }
}
