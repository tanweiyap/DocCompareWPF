using DocConvert;
using OpenCvSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocCompareDLL
{
    public class DocCompareClass
    {
        public static ArrayList DocCompare(ref string folder1, ref string folder2, ref string outfolder, ref int seqlen, int[,] force_pairs = null)
        {
            force_pairs ??= new int[0, 0]; //{ { 5, 5 }, { 6, 6 }, { 7, 7 }, { 8, 8 } };

            ArrayList doc1 = new ArrayList(); //all images of fist document
            ArrayList doc2 = new ArrayList(); //all images of second document

            //initialize vectors for output sequences
            ArrayList seqi = new ArrayList();
            ArrayList seqj = new ArrayList();
            ArrayList alignment = new ArrayList(); //to be returned... it is a concatenation of seqi and seqj
            ArrayList sizes = new ArrayList(); //individual image sizes of output document, for overlay

            //find jpgs in folder: only runs on Windows!!
            string[] fna = System.IO.Directory.GetFiles(folder1, "*.jpg");
            string[] fnb = System.IO.Directory.GetFiles(folder2, "*.jpg");
            string[] fnau = System.IO.Directory.GetFiles(folder1, "*.JPG");
            string[] fnbu = System.IO.Directory.GetFiles(folder2, "*.JPG");

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
                m++; //beacuse we start images at 0.jpg, there is one more image than the max.nr, e.g. 14.jpg is max --> in total 15 images.

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
                    Mat a = Cv2.ImRead(Path.Join(folder1, ii.ToString() + ".jpg"));
                    Cv2.Resize(a, a, work_size);
                    doc1.Add(a);
                }

                //load all images in folder 2
                for (int ii = 0; ii < n; ii++)
                {
                    Mat b = Cv2.ImRead(Path.Join(folder2, ii.ToString() + ".jpg"));
                    sizes.Add(b.Size());
                    Cv2.Resize(b, b, work_size);
                    doc2.Add(b);
                }

                double[,] distanceMatrix = new double[m, n];

                //fill DistanceMatrix
                for (int ii = 0; ii < m; ii++)
                {
                    for (int jj = 0; jj < n; jj++)
                    {
                        Match_score_jpg((Mat)doc1[ii], (Mat)doc2[jj], ref distanceMatrix[ii, jj]);
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
                    distanceMatrix[force_pairs[i, 0], force_pairs[i, 1]] = 100;
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
                        if (distanceMatrix[(int)seqi[i], (int)seqj[i]] != 1)
                        {
                            Mat difference = new Mat();
                            Image_compare((Mat)doc1[(int)seqi[i]], (Mat)doc2[(int)seqj[i]], (Size)sizes[(int)seqj[i]], ref difference);
                            Cv2.ImWrite(Path.Join(outfolder, (seqi[i]).ToString() + "_" + (seqj[i]).ToString() + ".png"), difference);
                            //namedWindow("image", WINDOW_NORMAL);
                            //imshow("image", difference);
                            //waitKey(0);
                        }
                    }
                }
            }
            return alignment;
        }

        // creates an image mask for the differences between two images
        private static void Image_compare(Mat alpha, Mat beta, Size orig_size, ref Mat diffHighlights)
        {
            Mat a = new Mat();
            Size my_size = new Size(512, 512);
            Cv2.Resize(alpha, a, my_size);
            Mat b = new Mat();
            Cv2.Resize(beta, b, my_size);
            Mat diff_reduced = new Mat();
            Size blur_size = new Size(10, 10);
            Cv2.Blur(a, a, blur_size);
            Cv2.Blur(b, b, blur_size);

            //generate shifted copies: 1 pixel in each direction

            Cv2.Absdiff(a, b, diff_reduced);
            //Cv2.NamedWindow("image", WindowMode.Normal);
            //Cv2.ImShow("image", diff_reduced);
            //Cv2.WaitKey(0);
            //convert to black&white
            Cv2.CvtColor(diff_reduced, diff_reduced, ColorConversionCodes.BGR2GRAY);
            Cv2.Threshold(diff_reduced, diff_reduced, 0, 255, ThresholdTypes.Binary);

            //dilate mask to remove pixel holes
            Cv2.Dilate(diff_reduced, diff_reduced, new Mat(), iterations: 3);
            //draw bounding boxes
            //Mat[] channels = new Mat[4];
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

            Mat al_ch = channels[0] * 0.7;

            Mat[] input = { blau, gruen, rot, al_ch };
            Mat[] output = { a };

            int[] from_to = {
                                0, 0, //zeros to channel B
								1, 1, //zeros to channel G
								2, 2, //ones to channel R
								3, 3  //contours to channel Alpha
			};
            Cv2.MixChannels(input, output, from_to);
            //finally, to output...
            Cv2.Resize(output[0], diffHighlights, orig_size);
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

        public static ArrayList TextCompare(ref TextDocumentClass doc1, ref TextDocumentClass doc2, ref int seqlen, out List<List<Diff>> diffList, int[,] force_pairs = null)
        {
            force_pairs ??= new int[0, 0];
            ArrayList alignment = new ArrayList();

            diff_match_patch diffMatch = new diff_match_patch();

            List<int> assignmentDoc1 = new List<int>(doc1.Paragraphs.Count);
            List<int> distDoc1 = new List<int>(doc1.Paragraphs.Count);
            List<List<Diff>> diffListDoc1 = new List<List<Diff>>();
            List<int> assignmentDoc2 = new List<int>(doc2.Paragraphs.Count);
            List<int> distDoc2 = new List<int>(doc2.Paragraphs.Count);
            List<List<Diff>> diffListDoc2 = new List<List<Diff>>();

            // init
            for (int i = 0; i < doc1.Paragraphs.Count; i++)
            {
                assignmentDoc1.Add(-1);
                distDoc1.Add(int.MaxValue);
                diffListDoc1.Add(new List<Diff>());
            }

            for (int i = 0; i < doc2.Paragraphs.Count; i++)
            {
                assignmentDoc2.Add(-1);
                distDoc2.Add(int.MaxValue);
                diffListDoc2.Add(new List<Diff>());
            }

            for (int i = 0; i < doc1.Paragraphs.Count; i++)
            {
                List<int> LocalDistance = new List<int>();
                List<Diff> bestList = new List<Diff>(); ;
                string selfText = doc1.Paragraphs[i].ToString();

                for (int j = 0; j < doc2.Paragraphs.Count; j++)
                {
                    string compText = doc2.Paragraphs[j].ToString();

                    List<Diff> DiffList = diffMatch.diff_main(selfText, compText);
                    diffMatch.diff_cleanupSemantic(DiffList);

                    int dist = diffMatch.diff_levenshtein(DiffList);
                    if (LocalDistance.Count != 0)
                    {
                        if (dist < LocalDistance.Min())
                        {
                            bestList = DiffList;
                        }
                    }
                    else
                    {
                        bestList = DiffList;
                    }
                    LocalDistance.Add(dist);
                }

                int indDistMin = LocalDistance.IndexOf(LocalDistance.Min());

                if (distDoc2[indDistMin] > LocalDistance[indDistMin] && LocalDistance[indDistMin] < Math.Max(selfText.Length, doc2.Paragraphs[indDistMin].ToString().Length) * 0.5) // if difference larger than 50 %
                {
                    if (assignmentDoc1.Contains(indDistMin))
                    {
                        if (distDoc1[assignmentDoc1.IndexOf(indDistMin)] > LocalDistance.Min())
                        {
                            diffListDoc1[assignmentDoc1.IndexOf(indDistMin)] = new List<Diff>();
                            assignmentDoc1[assignmentDoc1.IndexOf(indDistMin)] = -1;
                        }
                    }

                    assignmentDoc2[indDistMin] = i;
                    assignmentDoc1[i] = indDistMin;
                    distDoc1[i] = LocalDistance.Min();
                    distDoc2[indDistMin] = LocalDistance.Min();
                    diffListDoc1[i] = bestList;
                    diffListDoc2[indDistMin] = bestList;
                }
            }

            // index of non -1 (paragraph assigned)
            List<int> doc1Assigned = new List<int>();
            foreach (int i in assignmentDoc1)
            {
                if (i != -1)
                    doc1Assigned.Add(i);
            }

            doc1Assigned.Sort();
            List<int> indSortedDoc1 = new List<int>();
            foreach (int i in doc1Assigned)
            {
                indSortedDoc1.Add(assignmentDoc1.IndexOf(i));
            }

            List<int> doc2Assigned = new List<int>();
            foreach (int i in assignmentDoc2)
            {
                if (i != -1)
                    doc2Assigned.Add(i);
            }

            doc2Assigned.Sort();
            List<int> indSortedDoc2 = new List<int>();
            foreach (int i in doc2Assigned)
            {
                indSortedDoc2.Add(assignmentDoc2.IndexOf(i));
            }
            // sort

            List<int> seq1 = new List<int>();
            List<int> seq2 = new List<int>();
            int z = -2;

            for (int i = 0; i < Math.Max(doc1.Paragraphs.Count, doc2.Paragraphs.Count); i++)
            {
                // doc1 before doc2
                if (i < doc1.Paragraphs.Count)
                {
                    if (assignmentDoc1[i] == -1)
                    {
                        seq1.Add(i);
                        seq2.Add(-1);
                    }
                    // check if index is in indSorted
                    else if (indSortedDoc1.Contains(i))
                    {
                        int ind = indSortedDoc1.IndexOf(i);
                        // if ind == 0
                        if (ind == 0)
                        {
                            seq1.Add(i);
                            seq2.Add(assignmentDoc1[i]);
                        }
                        else if (indSortedDoc1[ind] > indSortedDoc1[ind - 1])
                        {
                            seq1.Add(i);
                            //seq2.Add(assignmentDoc1[i]);
                            seq2.Add(assignmentDoc1[indSortedDoc1[ind]]);
                        }
                        else
                        {
                            if (assignmentDoc1[i] > -1)
                            {
                                seq1.Add(i);
                                seq2.Add(z);
                                assignmentDoc1[i] = z;
                                assignmentDoc2[assignmentDoc2.IndexOf(i)] = z;
                                z--;
                            }
                            else
                            {
                                seq1.Add(i);
                                seq2.Add(assignmentDoc2[assignmentDoc2.IndexOf(assignmentDoc1[i])]);
                            }
                        }
                    }
                }

                // doc1 before doc2
                if (i < doc2.Paragraphs.Count)
                {
                    if (assignmentDoc2[i] == -1)
                    {
                        seq1.Add(-1);
                        seq2.Add(i);
                    }
                    // check if index is in indSorted
                    else if (indSortedDoc2.Contains(i))
                    {
                        int ind = indSortedDoc2.IndexOf(i);
                        // if ind == 0
                        if (ind == 0)
                        {
                            seq1.Add(assignmentDoc2[i]);
                            seq2.Add(i);
                        }
                        else if (indSortedDoc2[ind] > indSortedDoc2[ind - 1])
                        {
                            //seq1.Add(assignmentDoc1[i]);
                            seq1.Add(assignmentDoc2[indSortedDoc2[ind]]);
                            seq2.Add(i);
                        }
                        else
                        {
                            if (assignmentDoc2[i] > -1)
                            {
                                seq1.Add(z);
                                seq2.Add(i);
                                assignmentDoc1[assignmentDoc1.IndexOf(i)] = z;
                                assignmentDoc2[i] = z;

                                /*
                                if (seq1.Contains(i))
                                {
                                    seq2[seq1.IndexOf(i)] = z;
                                }
                                */
                                z--;
                            }
                            else
                            {
                                seq1.Add(assignmentDoc1[assignmentDoc1.IndexOf(assignmentDoc2[i])]);
                                seq2.Add(i);
                            }
                        }
                    }
                }
            }

            // clean up

            // find duplicate
            List<int> indToRemove = new List<int>();
            for (int i = 1; i < seq1.Count; i++)
            {
                if (seq1[i] == seq1[i - 1] && seq2[i] == seq2[i - 1])
                    indToRemove.Add(i);
            }

            foreach (int ind in indToRemove.OrderByDescending(v => v))
            {
                seq1.RemoveAt(ind);
                seq2.RemoveAt(ind);
            }

            // find cross similarity

            indToRemove.Clear();
            for (int i = 1; i < seq1.Count; i++)
            {
                if (seq1[i] == seq2[i - 1] && seq2[i] == seq1[i - 1] && seq1[i] != -1 && seq1[i - 1] != -1 && seq2[i] != -1 && seq2[i - 1] != -1)
                {
                    indToRemove.Add(i);
                    seq1[i] = seq2[i];
                    seq2[i - 1] = seq1[i - 1];
                }
            }

            foreach (int ind in indToRemove.OrderByDescending(v => v))
            {
                seq1.RemoveAt(ind);
                seq2.RemoveAt(ind);
            }

            // clean unnecessary indices < -1
            indToRemove.Clear();
            for (int i = 1; i < seq1.Count; i++)
            {
                if (seq1[i-1] == seq2[i] && seq1[i-1] < -1 && seq2[i] < -1)
                {
                    indToRemove.Add(i-1);
                    seq2[i] = seq2[i-1];                    
                }
            }

            foreach (int ind in indToRemove.OrderByDescending(v => v))
            {
                seq1.RemoveAt(ind);
                seq2.RemoveAt(ind);
            }
            List<int> seq11 = new List<int>();
            List<int> seq22 = new List<int>();

            for (int i = 0; i < seq1.Count; i++)
            {
                if (seq11.Contains(seq1[i]))
                {
                    if (seq22[seq11.IndexOf(seq1[i])] == seq2[i])
                    {
                    }
                    else
                    {
                        seq11.Add(seq1[i]);
                        seq22.Add(seq2[i]);
                    }
                }
                else
                {
                    seq11.Add(seq1[i]);
                    seq22.Add(seq2[i]);
                }
            }

            // index of non -1 (paragraph assigned)

            List<int> seq222 = new List<int>();
            foreach (int i in seq22)
            {
                if (i > -1)
                    seq222.Add(i);
            }

            seq222.Sort();
            List<int> indSortedSeq2 = new List<int>();
            foreach (int i in seq222)
            {
                indSortedSeq2.Add(seq22.IndexOf(i));
            }

            List<List<int>> indPairsToSwap = new List<List<int>>();
            for (int i = 1; i < indSortedSeq2.Count; i++)
            {
                if (indSortedSeq2[i] < indSortedSeq2[i - 1])
                {
                    int ind1 = seq22.IndexOf(seq22[indSortedSeq2[i - 1]]);
                    int ind2 = seq22.IndexOf(seq22[indSortedSeq2[i]]);
                    indPairsToSwap.Add(new List<int>() { ind1, ind2 });
                }
            }

            foreach (List<int> pair in indPairsToSwap)
            {
                int seq1Ind1 = seq11[pair[0]];
                int seq1Ind2 = seq11[pair[1]];

                seq11[pair[0]] = seq1Ind2;
                seq11[pair[1]] = seq1Ind1;

                int seq2Ind1 = seq22[pair[0]];
                int seq2Ind2 = seq22[pair[1]];

                seq22[pair[0]] = seq2Ind2;
                seq22[pair[1]] = seq2Ind1;
            }

            // create alignment and difflist
            for (int i = 0; i < seq11.Count; i++)
            {
                alignment.Add(seq11[i]);
            }

            for (int i = 0; i < seq22.Count; i++)
            {
                alignment.Add(seq22[i]);
            }

            diffList = new List<List<Diff>>();

            for (int i = 0; i < seq11.Count; i++)
            {
                if (seq11[i] >= 0)
                {
                    diffList.Add(diffListDoc1[seq11[i]]);
                }
                else
                {
                    diffList.Add(diffListDoc2[seq22[i]]);
                }
            }

            seqlen = diffList.Count;

            return alignment;
        }
    }
}