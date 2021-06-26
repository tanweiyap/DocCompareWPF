using OpenCvSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace DocCompareDLL
{
    public class DocCompareClass
    {
        public static ArrayList DocCompare(ref string folder1, ref string folder2, ref string outfolder, ref string outfolder2, ref int seqlen, int[,] force_pairs = null, List<List<int>> noCompareZones = null)
        {
            force_pairs ??= new int[0, 0]; //{ { 5, 5 }, { 6, 6 }, { 7, 7 }, { 8, 8 } };
            noCompareZones ??= new List<List<int>>();

            ArrayList doc1 = new ArrayList(); //all images of fist document
            ArrayList doc2 = new ArrayList(); //all images of second document

            //initialize vectors for output sequences
            ArrayList seqi = new ArrayList();
            ArrayList seqj = new ArrayList();
            ArrayList alignment = new ArrayList(); //to be returned... it is a concatenation of seqi and seqj
            ArrayList sizes = new ArrayList(); //individual image sizes of output document, for overlay

            //find jpgs in folder: only runs on Windows!!
            string[] fna = System.IO.Directory.GetFiles(folder1, "*.png");
            string[] fnb = System.IO.Directory.GetFiles(folder2, "*.png");
            string[] fnau = System.IO.Directory.GetFiles(folder1, "*.PNG");
            string[] fnbu = System.IO.Directory.GetFiles(folder2, "*.PNG");

            string[] fn1 = new string[fna.Length + fnau.Length];
            fna.CopyTo(fn1, 0);
            fnau.CopyTo(fn1, fna.Length);
            string[] fn2 = new string[fnb.Length + fnbu.Length];
            fnb.CopyTo(fn2, 0);
            fnbu.CopyTo(fn2, fnb.Length);

            Size work_size = new Size(512, 512);

            if (fn1.Length == 0 || fn2.Length == 0)
            {
                Console.WriteLine("one of the documents has no pages...");
            }
            else
            {
                //find the largest number image...
                int m = 0; //number of pages in doc1
                int n = 0; //number of pages in doc2

                //document1
                for (int i = 0; i < fn1.Length; i++)
                {
                    int pagenr1 = int.Parse(Path.GetFileNameWithoutExtension(fn1[i]));
                    if (m < pagenr1)
                    {
                        m = pagenr1;
                    }
                }
                m++; //beacuse we start images at 0.png, there is one more image than the max.nr, e.g. 14.png is max --> in total 15 images.

                //document2
                for (int i = 0; i < fn2.Length; i++)
                {
                    int pagenr2 = int.Parse(Path.GetFileNameWithoutExtension(fn2[i]));
                    if (n < pagenr2)
                    {
                        n = pagenr2;
                    }
                }
                n++;

                //load all images in folder 1
                for (int ii = 0; ii < m; ii++)
                {
                    Mat a = Cv2.ImRead(Path.Join(folder1, ii.ToString() + ".png"));
                    //Cv2.Resize(a, a, work_size);
                    doc1.Add(a);
                }

                //load all images in folder 2
                for (int ii = 0; ii < n; ii++)
                {
                    Mat b = Cv2.ImRead(Path.Join(folder2, ii.ToString() + ".png"));
                    sizes.Add(b.Size());
                    //Cv2.Resize(b, b, work_size);
                    doc2.Add(b);
                }

                double[,] distanceMatrix = new double[m, n];

                //fill DistanceMatrix
                for (int ii = 0; ii < m; ii++)
                {
                    for (int jj = 0; jj < n; jj++)
                    {
                        Mat c = new Mat();
                        Mat d = new Mat();
                        Cv2.Resize((Mat)doc1[ii], c, work_size);
                        Cv2.Resize((Mat)doc2[jj], d, work_size);
                        //Match_score_jpg((Mat)doc1[ii], (Mat)doc2[jj], ref distanceMatrix[ii, jj]);
                        Match_score_jpg(c, d, ref distanceMatrix[ii, jj]);
                    }
                }

                //normalize Distance Matrix
                double min_dist = 1.0;
                double max_dist = 0.0;
                //sum over all elements in matrix, store min and max value on the go
                for (int i = 0; i < m; i++)
                {
                    for (int j = 0; j < n; j++)
                    {
                        if (distanceMatrix[i, j] < min_dist)
                        {
                            min_dist = distanceMatrix[i, j];
                        }
                        if (distanceMatrix[i, j] > max_dist)
                        {
                            max_dist = distanceMatrix[i, j];
                        }
                    }
                }

                //get these differences scalet between 0 and 1 and inverted, so that 1 is the score for match and 0 the score for minimal match in this respective document
                //take all elements e in matrix 1 - (e - min) / (max - min)-- > transformed differences in scores
                //--> this is the matrix we will use.
                //- collect average non-1 scores
                int n_non1 = 0;
                double sum_non1 = 0;
                if (max_dist - min_dist > 0)
                {
                    for (int i = 0; i < m; i++)
                    {
                        for (int j = 0; j < n; j++)
                        {
                            if (distanceMatrix[i, j] != 0.0)
                            {
                                distanceMatrix[i, j] = (1 - ((distanceMatrix[i, j] - min_dist) / (max_dist - min_dist))) * 0.999999; //the last factor shall make sure that the smallest distance is recognized, because we skip '1' - distanceMatrix entries
                                sum_non1 += distanceMatrix[i, j];
                                n_non1++;
                            }
                            else
                            {
                                distanceMatrix[i, j] = 1;
                            }
                        }
                    }
                }

                double av_non1 = 1;
                //average score
                if (n_non1 > 0)
                {
                    av_non1 = sum_non1 / n_non1;
                }

                //else sum_non1 will be 0...
                double gap_penalty = 0.5 * av_non1;

                //catch single page, alignment should be tried without gap!
                if (n == 1 || m == 1)
                {
                    gap_penalty = -0.5;
                }

                //change distanceMatrix where fore_pairs is set
                for (int i = 0; i < (int)force_pairs.GetLength(0); i++)
                {
                    distanceMatrix[force_pairs[i, 0], force_pairs[i, 1]] = distanceMatrix[force_pairs[i, 0], force_pairs[i, 1]] + 100;
                }

                //distanceMatrix[5, 5] = 100;

                //create scoring matrix according to NMW algorithm.Directly
                //store the 'where from?' info: 0 : sequence i, 1 : sequence j, 2 : diagonal
                double[,] score = new double[m + 1, n + 1];
                int[,] pointers = new int[m, n];

                //first rowand column:
                for (int i = 0; i < m + 1; i++)
                {
                    score[i, 0] = gap_penalty * ((double)i + 1);
                }
                for (int j = 1; j < n + 1; j++)
                {
                    score[0, j] = gap_penalty * ((double)j + 1);
                }

                double match;
                double del;
                double insert;
                for (int i = 1; i < m + 1; i++)
                {
                    for (int j = 1; j < n + 1; j++)
                    {
                        match = score[i - 1, j - 1] + distanceMatrix[i - 1, j - 1];
                        del = score[i - 1, j] + gap_penalty;
                        insert = score[i, j - 1] + gap_penalty;

                        if (match > del)
                        {
                            if (match > insert)
                            {
                                score[i, j] = match;
                                pointers[i - 1, j - 1] = 2;
                            }
                            else
                            {
                                score[i, j] = insert;
                                pointers[i - 1, j - 1] = 0;
                            }
                        }
                        else
                        {
                            if (del > insert)
                            {
                                score[i, j] = del;
                                pointers[i - 1, j - 1] = 1;
                            }
                            else
                            {
                                score[i, j] = insert;
                                pointers[i - 1, j - 1] = 0;
                            }
                        }
                    }
                }

                //find maximal value to start with
                int currposi = m;
                int currposj = n;
                double mymax = score[currposi, currposj];

                for (int jj = 0; jj < n; jj++)
                {
                    if (score[currposi, jj] > mymax)
                    {
                        currposj = jj;
                        mymax = score[currposi, jj];
                    }
                }

                currposi -= 1;
                currposj -= 1;

                //in case alignment ends with shift
                for (int i = n - 1; i > currposj; i--)
                {
                    seqi.Add(-1);
                    seqj.Add(i);
                }

                //create the sequences
                while (currposi > -1 && currposj > -1)
                {
                    if (pointers[currposi, currposj] == 2)
                    {
                        seqi.Add(currposi);
                        seqj.Add(currposj);
                        currposi -= 1;
                        currposj -= 1;
                    }
                    else if (pointers[currposi, currposj] == 1)
                    {
                        seqi.Add(currposi);
                        seqj.Add(-1);
                        currposi -= 1;
                    }
                    else if (pointers[currposi, currposj] == 0)
                    {
                        seqi.Add(-1);
                        seqj.Add(currposj);
                        currposj -= 1;
                    }
                }

                //finalize strings if gaps exist at beginning.Only one of currposiand currposj will be non - zero
                for (int i = currposi; i > -1; i--)
                {
                    seqi.Add(i);
                    seqj.Add(-1);
                }
                for (int j = currposj; j > -1; j--)
                {
                    seqi.Add(-1);
                    seqj.Add(j);
                }

                //alignment = new ArrayList[2 * seqi.Count]; // needs to be deleted by calling routine!
                seqlen = seqi.Count;

                for (int ii = 0; ii < seqi.Count; ii++)
                {
                    alignment.Add(seqi[ii]);
                }
                for (int ii = 0; ii < seqi.Count; ii++)
                {
                    alignment.Add(seqj[ii]);
                }

                //for (int ii = 0; ii < 2 * seqi.Count; ii++)
                //{
                //	Console.Write(alignment[ii] + " ");
                //}
                //Console.WriteLine();

                //find out which images need an compare-overlay
                for (int i = 0; i < seqi.Count; i++)
                {
                    //only comparisons with real pages...
                    if ((int)seqi[i] != -1 && (int)seqj[i] != -1)
                    {
                        //only if they differ...
                        if ( ( distanceMatrix[(int)seqi[i], (int)seqj[i]] <= 0.9999 ) | ( distanceMatrix[(int)seqi[i], (int)seqj[i]] >= 100 & distanceMatrix[(int)seqi[i], (int)seqj[i]] < 100.9999) )
                        {
                            Mat differencered = new Mat();
                            Mat differencegreen = new Mat();
                            Image_compare((Mat)doc1[(int)seqi[i]], (Mat)doc2[(int)seqj[i]], (Size)sizes[(int)seqj[i]], ref differencered, ref differencegreen);
                            Cv2.ImWrite(Path.Join(outfolder, (seqi[i]).ToString() + "_" + (seqj[i]).ToString() + ".png"), differencered);
                            Cv2.ImWrite(Path.Join(outfolder2, (seqi[i]).ToString() + "_" + (seqj[i]).ToString() + ".png"), differencegreen);                            
                        }
                    }
                }
            }
            //alignment: two times n-1 ... 0 (backwards), -1 for gaps
            return alignment;
        }

        //public static ArrayList DocCompare(ref string folder1, ref string folder2, ref string outfolder, ref int seqlen, int[,] force_pairs = null)
        public static ArrayList DocCompareMult(ref string folder1, ref string folder2, int previous_docs, ArrayList previous_alignment = null)
        //previous alignment: {reference;comp1; comp2; ...}. previous_docs is the number of already compared docs WITHOUT THE REFERENCE - so can be 0. Then, previous alignment should be empty.
        {
            //ArrayList previous_alignment;
            //ArrayList previous_alignmentr;
            //ArrayList previous_alignmentc;
            //int previous_docs = 1;
            previous_alignment ??= new ArrayList();

            //previous_alignmentr = new ArrayList();
            //previous_alignmentr.Add(-1);
            //previous_alignmentr.Add(-1);
            //previous_alignmentr.Add(3);
            //previous_alignmentr.Add(2);
            //previous_alignmentr.Add(-1);
            //previous_alignmentr.Add(1);
            //previous_alignmentr.Add(0);
            //previous_alignmentr.Add(-1);

            //previous_alignmentc = new ArrayList();
            //previous_alignmentc.Add(-1);
            //previous_alignmentc.Add(4);
            //previous_alignmentc.Add(-1);
            //previous_alignmentc.Add(3);
            //previous_alignmentc.Add(2);
            //previous_alignmentc.Add(-1);
            //previous_alignmentc.Add(1);
            //previous_alignmentc.Add(0);

            //previous_alignment.Add(previous_alignmentr);
            //previous_alignment.Add(previous_alignmentc);




            ArrayList doc1 = new ArrayList(); //all images of fist document
            ArrayList doc2 = new ArrayList(); //all images of second document

            //initialize vectors for output sequences
            ArrayList seqi = new ArrayList();
            ArrayList seqj = new ArrayList();
            ArrayList alignment = new ArrayList(); //to be returned... it is a concatenation of seqi and seqj
            ArrayList sizes = new ArrayList(); //individual image sizes of output document, for overlay

            //find jpgs in folder: only runs on Windows!!
            string[] fna = System.IO.Directory.GetFiles(folder1, "*.png");
            string[] fnb = System.IO.Directory.GetFiles(folder2, "*.png");
            string[] fnau = System.IO.Directory.GetFiles(folder1, "*.PNG");
            string[] fnbu = System.IO.Directory.GetFiles(folder2, "*.PNG");

            string[] fn1 = new string[fna.Length + fnau.Length];
            fna.CopyTo(fn1, 0);
            fnau.CopyTo(fn1, fna.Length);
            string[] fn2 = new string[fnb.Length + fnbu.Length];
            fnb.CopyTo(fn2, 0);
            fnbu.CopyTo(fn2, fnb.Length);

            Size work_size = new Size(300, 300);

            if (fn1.Length == 0 || fn2.Length == 0)
            {
                Console.WriteLine("one of the documents has no pages...");
            }
            else
            {
                //find the largest number image...
                int m = 0; //number of pages in doc1
                int n = 0; //number of pages in doc2

                //document1
                for (int i = 0; i < fn1.Length; i++)
                {
                    int pagenr1 = int.Parse(Path.GetFileNameWithoutExtension(fn1[i]));
                    if (m < pagenr1)
                    {
                        m = pagenr1;
                    }
                }
                m++; //beacuse we start images at 0.png, there is one more image than the max.nr, e.g. 14.png is max --> in total 15 images.

                //document2
                for (int i = 0; i < fn2.Length; i++)
                {
                    int pagenr2 = int.Parse(Path.GetFileNameWithoutExtension(fn2[i]));
                    if (n < pagenr2)
                    {
                        n = pagenr2;
                    }
                }
                n++;

                //load all images in folder 1
                for (int ii = 0; ii < m; ii++)
                {
                    Mat a = Cv2.ImRead(Path.Join(folder1, ii.ToString() + ".png"));
                    //Cv2.Resize(a, a, work_size);
                    doc1.Add(a);
                }

                //load all images in folder 2
                for (int ii = 0; ii < n; ii++)
                {
                    Mat b = Cv2.ImRead(Path.Join(folder2, ii.ToString() + ".png"));
                    sizes.Add(b.Size());
                    //Cv2.Resize(b, b, work_size);
                    doc2.Add(b);
                }

                double[,] distanceMatrix = new double[m, n];

                //fill DistanceMatrix
                for (int ii = 0; ii < m; ii++)
                {
                    for (int jj = 0; jj < n; jj++)
                    {
                        Mat c = new Mat();
                        Mat d = new Mat();
                        Cv2.Resize((Mat)doc1[ii], c, work_size, interpolation: 0);
                        Cv2.Resize((Mat)doc2[jj], d, work_size, interpolation: 0);
                        //Match_score_jpg((Mat)doc1[ii], (Mat)doc2[jj], ref distanceMatrix[ii, jj]);
                        Match_score_jpg(c, d, ref distanceMatrix[ii, jj]);
                    }
                }

                //normalize Distance Matrix
                double min_dist = 1.0;
                double max_dist = 0.0;
                //sum over all elements in matrix, store min and max value on the go
                for (int i = 0; i < m; i++)
                {
                    for (int j = 0; j < n; j++)
                    {
                        if (distanceMatrix[i, j] < min_dist)
                        {
                            min_dist = distanceMatrix[i, j];
                        }
                        if (distanceMatrix[i, j] > max_dist)
                        {
                            max_dist = distanceMatrix[i, j];
                        }
                    }
                }

                //get these differences scalet between 0 and 1 and inverted, so that 1 is the score for match and 0 the score for minimal match in this respective document
                //take all elements e in matrix 1 - (e - min) / (max - min)-- > transformed differences in scores
                //--> this is the matrix we will use.
                //- collect average non-1 scores
                int n_non1 = 0;
                double sum_non1 = 0;
                if (max_dist - min_dist > 0)
                {
                    for (int i = 0; i < m; i++)
                    {
                        for (int j = 0; j < n; j++)
                        {
                            if (distanceMatrix[i, j] != 0.0)
                            {
                                distanceMatrix[i, j] = (1 - ((distanceMatrix[i, j] - min_dist) / (max_dist - min_dist))) * 0.999999; //the last factor shall make sure that the smallest distance is recognized, because we skip '1' - distanceMatrix entries
                                sum_non1 += distanceMatrix[i, j];
                                n_non1++;
                            }
                            else
                            {
                                distanceMatrix[i, j] = 1;
                            }
                        }
                    }
                }

                double av_non1 = 1;
                //average score
                if (n_non1 > 0)
                {
                    av_non1 = sum_non1 / n_non1;
                }

                //else sum_non1 will be 0...
                double gap_penalty = 0.5 * av_non1;

                //catch single page, alignment should be tried without gap!
                if (n == 1 || m == 1)
                {
                    gap_penalty = -0.5;
                }

                //create scoring matrix according to NMW algorithm.Directly
                //store the 'where from?' info: 0 : sequence i, 1 : sequence j, 2 : diagonal
                double[,] score = new double[m + 1, n + 1];
                int[,] pointers = new int[m, n];

                //first rowand column:
                for (int i = 0; i < m + 1; i++)
                {
                    score[i, 0] = gap_penalty * ((double)i + 1);
                }
                for (int j = 1; j < n + 1; j++)
                {
                    score[0, j] = gap_penalty * ((double)j + 1);
                }

                double match;
                double del;
                double insert;
                for (int i = 1; i < m + 1; i++)
                {
                    for (int j = 1; j < n + 1; j++)
                    {
                        match = score[i - 1, j - 1] + distanceMatrix[i - 1, j - 1];
                        del = score[i - 1, j] + gap_penalty;
                        insert = score[i, j - 1] + gap_penalty;

                        if (match > del)
                        {
                            if (match > insert)
                            {
                                score[i, j] = match;
                                pointers[i - 1, j - 1] = 2;
                            }
                            else
                            {
                                score[i, j] = insert;
                                pointers[i - 1, j - 1] = 0;
                            }
                        }
                        else
                        {
                            if (del > insert)
                            {
                                score[i, j] = del;
                                pointers[i - 1, j - 1] = 1;
                            }
                            else
                            {
                                score[i, j] = insert;
                                pointers[i - 1, j - 1] = 0;
                            }
                        }
                    }
                }

                //find maximal value to start with
                int currposi = m;
                int currposj = n;
                double mymax = score[currposi, currposj];

                for (int jj = 0; jj < n; jj++)
                {
                    if (score[currposi, jj] > mymax)
                    {
                        currposj = jj;
                        mymax = score[currposi, jj];
                    }
                }

                currposi -= 1;
                currposj -= 1;

                //in case alignment ends with shift
                for (int i = n - 1; i > currposj; i--)
                {
                    seqi.Add(-1);
                    seqj.Add(i);
                }

                //create the sequences
                while (currposi > -1 && currposj > -1)
                {
                    if (pointers[currposi, currposj] == 2)
                    {
                        seqi.Add(currposi);
                        seqj.Add(currposj);
                        currposi -= 1;
                        currposj -= 1;
                    }
                    else if (pointers[currposi, currposj] == 1)
                    {
                        seqi.Add(currposi);
                        seqj.Add(-1);
                        currposi -= 1;
                    }
                    else if (pointers[currposi, currposj] == 0)
                    {
                        seqi.Add(-1);
                        seqj.Add(currposj);
                        currposj -= 1;
                    }
                }

                //finalize strings if gaps exist at beginning.Only one of currposiand currposj will be non - zero
                for (int i = currposi; i > -1; i--)
                {
                    seqi.Add(i);
                    seqj.Add(-1);
                }
                for (int j = currposj; j > -1; j--)
                {
                    seqi.Add(-1);
                    seqj.Add(j);
                }
                if (previous_docs == 0)
                {
                    alignment.Add(seqi);
                    alignment.Add(seqj);


                    //for (int ii = 0; ii < 2 * seqi.Count; ii++)
                    //{
                    //	Console.Write(alignment[ii] + " ");
                    //}
                    //Console.WriteLine();


                    //check if previous alignment exists
                    //seqi.Add(-1);
                    //seqi.Add(3);
                    //seqi.Add(-1);
                    //seqi.Add(-1);
                    //seqi.Add(2);
                    //seqi.Add(1);
                    //seqi.Add(0);

                    //seqj.Add(3);
                    //seqj.Add(-1);
                    //seqj.Add(2);
                    //seqj.Add(1);
                    //seqj.Add(0);
                    //seqj.Add(-1);
                    //seqj.Add(-1);
                }
                else
                {
                    //create working copies and assists
                    ArrayList prevalig = (ArrayList)previous_alignment[0];
                    int old_len = prevalig.Count / (previous_docs + 1); //length of previous alignment

                    ArrayList newref = new ArrayList();
                    newref.AddRange(seqi);
                    ArrayList oldref = new ArrayList();
                    oldref.AddRange(prevalig);
                    ArrayList newcompared = new ArrayList();
                    newcompared.AddRange(seqj);
                    ArrayList oldcompared = new ArrayList();
                    oldcompared.AddRange(previous_alignment.GetRange(1, previous_docs));


                    int lastseen = -1;
                    int insertindex = newref.Count - 1;
                    //1. Shift the newcompared by all -1s in Reference and write to new array
                    for (int ii = oldref.Count - 1; ii > -1; ii--)
                    {

                        int curr_element = (int)oldref[ii];
                        //check if it is -1
                        if (curr_element != -1)
                        {
                            //write to last seen
                            lastseen = curr_element;
                        }
                        else
                        {
                            //find lastseen in new alignment reference
                            if (lastseen != -1)
                            {
                                insertindex = newref.IndexOf(lastseen);
                                lastseen = -1;
                            }
                            //shift everything in new alignment after that index by 1
                            //newcompared.Insert(insertindex + 1, -1);
                            newcompared.Insert(insertindex, -1);
                        }

                    }

                    //2. Shift all previous alignments by all -1s in new alignment first half
                    lastseen = -1;
                    insertindex = oldref.Count - 1;
                    //1. Shift the new alignment second half by all -1s in Reference and write to new array
                    for (int ii = newref.Count - 1; ii > -1; ii--)
                    {
                        int curr_element = (int)newref[ii];
                        //check if it is -1
                        if (curr_element != -1)
                        {
                            //write to last seen
                            lastseen = curr_element;
                        }
                        else
                        {
                            //find lastseen in new alignment reference

                            if (lastseen != -1)
                            {
                                insertindex = oldref.IndexOf(lastseen);
                                lastseen = -1;
                            }
                            //shift everything in old alignments after that index by 1
                            oldref.Insert(insertindex, -1);

                            foreach (ArrayList innerArr in oldcompared)
                            {
                                innerArr.Insert(insertindex, -1);
                            }

                        }

                    }
                    //3. Remove all empty rows
                    for (int ii = newcompared.Count - 1; ii > -1; ii--)
                    {
                        bool allminus = true;
                        if ((int)newcompared[ii] == -1 && (int)oldref[ii] == -1)
                        {
                            foreach (ArrayList innerArr in oldcompared)
                            {
                                if ((int)innerArr[ii] != -1)
                                {
                                    allminus = false;
                                }
                            }
                        }
                        else
                        {
                            allminus = false;
                        }
                        if (allminus)
                        {
                            foreach (ArrayList innerArr in oldcompared)
                            {
                                innerArr.RemoveAt(ii);
                            }
                            oldref.RemoveAt(ii);
                            newcompared.RemoveAt(ii);

                        }

                    }
                    //4. Unify subsequent "-1 in reference ii and ii-1"
                    for (int ii = newcompared.Count - 1; ii > 0; ii--) //0 because 0th element cannot be shifted upwards!
                    {
                        if ((int)oldref[ii] == -1 && (int)oldref[ii - 1] == -1)
                        {
                            bool possible = true;
                            //check for each element if above is -1
                            foreach (ArrayList innerArr in oldcompared)
                            {
                                if ((int)innerArr[ii] != -1)
                                {
                                    if ((int)innerArr[ii - 1] != -1)
                                    {
                                        possible = false;
                                    }
                                }
                            }
                            if (possible && (int)newcompared[ii] != -1)
                            {
                                if ((int)newcompared[ii - 1] != -1)
                                {
                                    possible = false;
                                }
                            }
                            if (possible)
                            {
                                //add 1 to all entries and add to previous line, then delete
                                foreach (ArrayList innerArr in oldcompared)
                                {
                                    innerArr[ii] = (int)innerArr[ii] + 1;
                                }
                                newcompared[ii] = (int)newcompared[ii] + 1;
                                //add to previous
                                foreach (ArrayList innerArr in oldcompared)
                                {
                                    innerArr[ii - 1] = (int)innerArr[ii - 1] + (int)innerArr[ii];
                                }
                                newcompared[ii - 1] = (int)newcompared[ii - 1] + (int)newcompared[ii];
                                //now remove line ii from oldref, oldcompared and newcompared
                                oldref.RemoveAt(ii);
                                newcompared.RemoveAt(ii);
                                foreach (ArrayList innerArr in oldcompared)
                                {
                                    innerArr.RemoveAt(ii);
                                }
                            }
                        }
                    }

                    //Compile return...
                    alignment.Add(oldref);
                    foreach (ArrayList innerArr in oldcompared)
                    {
                        alignment.Add(innerArr);
                    }
                    alignment.Add(newcompared);
                }
            }
            return alignment;
        }


        // creates an image mask for the differences between two images
        private static void Image_compare(Mat alpha, Mat beta, Size orig_size, ref Mat diffHighlights, ref Mat diffHighlightsG)
        {
            Mat a = new Mat();

            Size my_size = alpha.Size();
            //Size my_size = new Size(1024, 1024);
            Cv2.Resize(alpha, a, my_size);
            Mat b = new Mat();
            Cv2.Resize(beta, b, my_size);
            Mat diff_reduced = new Mat();
            Size blur_size = new Size(10, 10);
            //Cv2.Blur(a, a, blur_size);
            //Cv2.Blur(b, b, blur_size);

            Cv2.Absdiff(a, b, diff_reduced);

            //Cv2.NamedWindow("image", WindowMode.Normal);
            //Cv2.ImShow("image", diff_reduced);
            //Cv2.WaitKey(0);

            //Cv2.NamedWindow("image", WindowMode.Normal);
            //Cv2.ImShow("image", alpha);
            //Cv2.WaitKey(0);

            //convert to black&white
            Cv2.CvtColor(diff_reduced, diff_reduced, ColorConversionCodes.BGR2GRAY);
            Cv2.Threshold(diff_reduced, diff_reduced, 0, 255, ThresholdTypes.Binary);

            //dilate mask to remove pixel holes
            Cv2.Dilate(diff_reduced, diff_reduced, new Mat(), iterations: 2);
            Cv2.Erode(diff_reduced, diff_reduced, new Mat(), iterations: 3);
            Cv2.Dilate(diff_reduced, diff_reduced, new Mat(), iterations: 2);
            //draw bounding boxes:
            Cv2.Split(diff_reduced, out Mat[] channels);

            //find out if too much on page is different: i.e. more than hlaf of the pixels are marked...
            int ref_fill = 128 * my_size.Height * my_size.Width;
            if (Cv2.Sum(channels[0])[0] > ref_fill)
            {
                channels[0] = channels[0] * 0 + 255;    //fill completely
            }
            else
            {
                /// Detect edges using canny
                int lowThreshold = 0;
                int ratio = 3;
                int kernel_size = 5;
                Cv2.Canny(channels[0], channels[0], lowThreshold, lowThreshold * ratio, kernel_size);
                /// Find contours
                Point[][] contours = Cv2.FindContoursAsArray(channels[0], RetrievalModes.List, ContourApproximationModes.ApproxSimple);
                Scalar colr = new Scalar(255, 255, 255);
                for (int i = 0; i < contours.Length; i++)
                {
                    Rect bounding_rect = Cv2.BoundingRect(contours[i]);
                    Cv2.Rectangle(channels[0], bounding_rect.TopLeft, bounding_rect.BottomRight, colr, thickness: -1);
                }
            }

            Cv2.CvtColor(a, a, ColorConversionCodes.BGR2BGRA);
            //here, wemix the channels to get a semi-transparent, red overlay
            Mat rot = 0 * channels[0] + 255;
            Mat gruen = 0 * channels[0] + 44;
            Mat blau = 0 * channels[0] + 108;

            Mat rot2 = 0 * channels[0] + 70;
            Mat gruen2 = 0 * channels[0] + 255;
            Mat blau2 = 0 * channels[0] + 183;

            Mat al_ch = channels[0] * 0.7;

            Mat[] input1 = { blau, gruen, rot, al_ch };
            Mat[] input2 = { blau2, gruen2, rot2, al_ch };
            Mat[] output = { a };
            Mat[] output2 = { a.Clone() };

            int[] from_to = {
                                0, 0, //zeros to channel B
								1, 1, //zeros to channel G
								2, 2, //ones to channel R
								3, 3  //contours to channel Alpha
			};
            Cv2.MixChannels(input1, output, from_to);
            Cv2.MixChannels(input2, output2, from_to);
            //finally, to output...
            Cv2.Resize(output[0], diffHighlights, orig_size);
            Cv2.Resize(output2[0], diffHighlightsG, orig_size);
        }

        //calculates the match-score for two images
        private static void Match_score_jpg(Mat alpha, Mat beta, ref double diff)
        {
            Mat gamma = new Mat();
            Cv2.Resize(beta, gamma, alpha.Size());
            Mat diffImg = new Mat();
            Cv2.Absdiff(alpha, gamma, diffImg);
            Cv2.CvtColor(diffImg, diffImg, ColorConversionCodes.BGR2GRAY);

            //normalize to 1 for black/white difference and 0 for exact same image...
            diff = (Cv2.Sum(diffImg)[0] + Cv2.Sum(diffImg)[1] + Cv2.Sum(diffImg)[2]) / (3 * 255 * (int)alpha.Rows * (int)alpha.Cols);
        }

        //get similarity score for two strings
        private static int Needle_wunsch(string refSeq, string alignSeq)
        {
            int max_score = 0;
            int refSeqCnt = refSeq.Length + 1;
            int alineSeqCnt = alignSeq.Length + 1;

            int[,] scoringMatrix = new int[alineSeqCnt, refSeqCnt];

            //Initialization Step - filled with 0 for the first row and the first column of matrix
            for (int i = 0; i < alineSeqCnt; i++)
            {
                scoringMatrix[i, 0] = 0;
            }

            for (int j = 0; j < refSeqCnt; j++)
            {
                scoringMatrix[0, j] = 0;
            }

            for (int i = 1; i < alineSeqCnt; i++)
            {
                for (int j = 1; j < refSeqCnt; j++)
                {
                    int scoreDiag;
                    if (refSeq.Substring(j - 1, 1) == alignSeq.Substring(i - 1, 1))
                        scoreDiag = scoringMatrix[i - 1, j - 1] + 2;
                    else
                        scoreDiag = scoringMatrix[i - 1, j - 1] + -1;

                    int scroeLeft = scoringMatrix[i, j - 1] - 2;
                    int scroeUp = scoringMatrix[i - 1, j] - 2;

                    int maxScore = Math.Max(Math.Max(scoreDiag, scroeLeft), scroeUp);

                    scoringMatrix[i, j] = maxScore;
                    if (maxScore > max_score) //collecting for final output
                    {
                        max_score = maxScore;
                    }
                }
            }

            //Traceback Step
            //char[] alineSeqArray = alignSeq.ToCharArray();
            //char[] refSeqArray = refSeq.ToCharArray();

            //string AlignmentA = string.Empty;
            //string AlignmentB = string.Empty;
            //int m = alineSeqCnt - 1;
            //int n = refSeqCnt - 1;
            //while (m > 0 && n > 0)
            //{
            //	int scroeDiag = 0;

            //	//Remembering that the scoring scheme is +2 for a match, -1 for a mismatch and -2 for a gap
            //	if (alineSeqArray[m - 1] == refSeqArray[n - 1])
            //		scroeDiag = 2;
            //	else
            //		scroeDiag = -1;

            //	if (m > 0 && n > 0 && scoringMatrix[m, n] == scoringMatrix[m - 1, n - 1] + scroeDiag)
            //	{
            //		AlignmentA = refSeqArray[n - 1] + AlignmentA;
            //		AlignmentB = alineSeqArray[m - 1] + AlignmentB;
            //		m = m - 1;
            //		n = n - 1;
            //	}
            //	else if (n > 0 && scoringMatrix[m, n] == scoringMatrix[m, n - 1] - 2)
            //	{
            //		AlignmentA = refSeqArray[n - 1] + AlignmentA;
            //		AlignmentB = "-" + AlignmentB;
            //		n = n - 1;
            //	}
            //	else //if (m > 0 && scoringMatrix[m, n] == scoringMatrix[m - 1, n] + -2)
            //	{
            //		AlignmentA = "-" + AlignmentA;
            //		AlignmentB = alineSeqArray[m - 1] + AlignmentB;
            //		m = m - 1;
            //	}
            //}
            return max_score;
        }

        //returns length of alignment. alignment will store the alignment and is size: (2* length of alignment)
    }
}