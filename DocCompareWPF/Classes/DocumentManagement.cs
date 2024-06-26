﻿using DocCompareDLL;
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
        public List<List<int>> noCompareZones;
        public int MAX_DOC_COUNT = 5;
        public ArrayList pageCompareIndices;
        public int totalLen;
        public string workingDir;
        public List<List<Diff>> pptSpeakerNotesDiff;
        public ArrayList globalAlignment;
        public bool doneGlobalAlignment;

        public DocumentManagement(string p_workingDir, AppSettings settings)
        {
            documents = new List<Document>();
            // TODO: Premium
            /*
            if (settings.numPanelsDragDrop == 3)
                documentsToShow = new List<int>() { 0, 1, 2 };
            else
                documentsToShow = new List<int>() { 0, 1 };
            */

            documentsToShow = new List<int>() { 0, 1, 2, 3, 4 };
            documentsToCompare = new List<int>() { 0, 1 };
            workingDir = p_workingDir;

            Directory.CreateDirectory(workingDir);

            Directory.CreateDirectory(Path.Join(workingDir, "compare"));
            forceAlignmentIndices = new List<List<int>>();
            noCompareZones = new List<List<int>>();
        }

        public DocumentManagement(int p_maxDocCount, string p_workingDir, AppSettings settings)
        {
            MAX_DOC_COUNT = p_maxDocCount;
            documents = new List<Document>();
            // TODO: Premium
            /*
            if (settings.numPanelsDragDrop == 3)
                documentsToShow = new List<int>() { 0, 1, 2 };
            else
                documentsToShow = new List<int>() { 0, 1 };
            */

            documentsToShow = new List<int>() { 0, 1, 2, 3, 4 };
            documentsToCompare = new List<int>() { 0, 1 };
            workingDir = p_workingDir;

            _ = Directory.CreateDirectory(workingDir);
            _ = new DirectoryInfo(workingDir)
            {
                Attributes = FileAttributes.Directory | FileAttributes.Hidden
            };

            forceAlignmentIndices = new List<List<int>>();
            noCompareZones = new List<List<int>>();
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
            // TODO: Premium
            /*
            List<string> docShown = new List<string>();
            for (int i = 0; i < documentsToShow.Count; i++)
            {
                if (documentsToShow[i] != -1 && i < documents.Count)
                {
                    docShown.Add(documents[documentsToShow[i]].docID);
                }
            }
            */
            // clean up folder

            DirectoryInfo di = new DirectoryInfo(documents[index].imageFolder);
            di.Delete(true);
            documents.RemoveAt(index);

            // shift everything down
            for (int i = viewID; i < documentsToShow.Count - 1; i++)
            {
                documentsToShow[i] = documentsToShow[i + 1];
            }

            documentsToShow[^1] = -1;

            for (int i = 0; i < documentsToShow.Count; i++)
            {
                if (documentsToShow[i] > index)
                {
                    documentsToShow[i]--;
                }
            }
            /*
            // we now check for docs to show
            for (int i = 0; i < docShown.Count - 1; i++)
            {
                if (i == viewID) // view, where the doc was removed
                {
                    string docIDToConsider;
                    for (int j = 0; j < documents.Count - 1; j++)
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
                documentsToShow[1] = -1;
            }

            for (int i = 0; i < documentsToShow.Count - 1; i++)
            {
                if (documentsToShow[i] == -1)
                {
                    documentsToShow.RemoveAt(i);
                    documentsToShow.Add(-1);
                }
            }

            if (documents.Count > 1)
            {
                for (int i = 0; i < documentsToShow.Count - 1; i++)
                {
                    if (documentsToShow[i] > index)
                    {
                        documentsToShow[i]--;
                    }
                }
            }
            */

            // default for doc comparison
            documentsToCompare[0] = -1;
            documentsToCompare[1] = -1;
        }

        public void RemoveDocumentWithID(string ID)
        {
            int ind = -1;
            for (int i = 0; i < documents.Count; i++)
            {
                if (documents[i].docID == ID)
                {
                    ind = i;
                    break;
                }
            }

            if (ind != -1)
            {
                List<string> docShown = new List<string>();
                for (int i = 0; i < documentsToShow.Count; i++)
                {
                    if (documentsToShow[i] != -1 && i < documents.Count)
                    {
                        docShown.Add(documents[documentsToShow[i]].docID);
                    }
                }

                DirectoryInfo di = new DirectoryInfo(documents[ind].imageFolder);

                if (di.Exists)
                {
                    di.Delete(true);
                }

                documents.RemoveAt(ind);

                for (int i = 0; i < docShown.Count; i++)
                {
                    if (i == 0) // view, where the doc was removed
                    {
                        string docIDToConsider;
                        for (int j = 0; j < documents.Count; j++)
                        {
                            docIDToConsider = documents[j].docID;
                            bool ok = false;
                            for (int k = 0; k < docShown.Count; k++)
                            {
                                if (k != 0)
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
                    documentsToShow[1] = -1;
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