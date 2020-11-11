using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace DocCompareWPF.Classes
{
    internal class DocumentManagement
    {
        public int displayToReload;
        public int docToReload;
        public List<Document> documents;
        public List<int> documentsToCompare;
        public List<int> documentsToShow;
        public List<List<int>> forceAlignmentIndices;
        public int MAX_DOC_COUNT = 5;
        public ArrayList pageCompareIndices;
        public int totalLen;
        public string workingDir;

        public DocumentManagement(string p_workingDir, AppSettings settings)
        {
            documents = new List<Document>();
            if (settings.numPanelsDragDrop == 3)
                documentsToShow = new List<int>() { 0, 1, 2 };
            else
                documentsToShow = new List<int>() { 0, 1 };

            documentsToCompare = new List<int>() { 0, 1 };
            workingDir = p_workingDir;

            DirectoryInfo di = new DirectoryInfo(workingDir);
            if (di.Exists == true)
                di.Delete(true);

            Directory.CreateDirectory(workingDir);

            Directory.CreateDirectory(Path.Join(workingDir, "compare"));
            forceAlignmentIndices = new List<List<int>>();
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
            if (di.Exists == true)
                di.Delete(true);
            _ = Directory.CreateDirectory(workingDir);
            _ = new DirectoryInfo(workingDir)
            {
                Attributes = FileAttributes.Directory | FileAttributes.Hidden
            };

            _ = Directory.CreateDirectory(Path.Join(workingDir, "compare"));
            forceAlignmentIndices = new List<List<int>>();
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
            if (documents.Count != 0 && documents.Count <= documentsToShow.Count)
            {
                if (documentsToShow[documents.Count - 1] == -1)
                {
                    documentsToShow[documents.Count - 1] = documents.Count - 1;
                }
            }
        }

        public void AddForceAligmentPairs(int source, int target)
        {
            forceAlignmentIndices.Add(new List<int>() { source, target });
        }

        public void RemoveDocument(int index, int viewID)
        {
            //string[] docShown = new string[3] { documents[documentsToShow[0]].docID, documents[documentsToShow[1]].docID, documents[documentsToShow[2]].docID };
            List<string> docShown = new List<string>();
            for (int i = 0; i < documentsToShow.Count; i++)
            {
                if (documentsToShow[i] != -1 && i < documents.Count)
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
            for (int i = 0; i < docShown.Count; i++)
            {
                if (i == viewID) // view, where the doc was removed
                {
                    string docIDToConsider;
                    for (int j = 0; j < documents.Count; j++)
                    {
                        docIDToConsider = documents[j].docID;
                        bool ok = false;
                        for (int k = 0; k < docShown.Count; k++)
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

            if (documents.Count == 0)
            {
                documentsToShow[0] = -1;
            }

            for (int i = 0; i < documentsToShow.Count; i++)
            {
                if (documentsToShow[i] == -1)
                {
                    documentsToShow.RemoveAt(i);
                    documentsToShow.Add(-1);
                }
            }
        }
        public void RemoveForceAligmentPairs(int source)
        {
            int index = -1;
            for (int i = 0; i < forceAlignmentIndices.Count; i++)
            {
                if (forceAlignmentIndices[i][0] == source)
                {
                    index = i;
                    break;
                }
            }

            if (index != -1)
                forceAlignmentIndices.RemoveAt(index);
        }
    }
}