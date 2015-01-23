using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace VisualTexcelProcessor
{
    class Processor
    {
        //excel application object
        private static Application excel = new Application();
        //Workbook object- contains sheets
        private static Workbook workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        //Aggregate BMP image, has generated BMPs added to it which are later overlaid to make a heatmap
        private List<Bitmap> bmpAggregate = new List<Bitmap>();
        //string that will hold the subject file name
        private string SubjectFile = "";
        /// <summary>
        /// Returns Task<bool>success</bool>, runs async
        /// </summary>
        /// <param name="subjectNumber"></param>
        /// <returns></returns>
        public async Task<bool> FileProcessor(int subjectNumber)
        {
            SubjectFile = "subject" + subjectNumber + ".txt";
            StreamReader sr = new StreamReader(SubjectFile);

            //I *think* this should handle missing files without halting the program
            //ex. 1 2 3 5 6 7 on a sequential 1-7 should just skip 4 and keep going
            //**UNTESTED**x
            if (sr == null) { Console.Write("Streamreader failed"); return false; }

            string raw = sr.ReadToEnd();
            string[] sample = raw.Split(';');

            Bitmap NewImage = await ProcessFile(sample, subjectNumber);
            bmpAggregate.Add(NewImage);

            return true;
        }
        /// <summary>
        /// Returns Task<Bitmap>pathTaken</Bitmap>, runs async, called by Processor for each workbook page
        /// </summary>
        /// <param name="sample"></param>
        /// <param name="subjectNumber"></param>
        /// <returns></returns>
        private Task<Bitmap> ProcessFile(string[] sample, int subjectNumber)
        {
            Worksheet ws = workbook.Worksheets.Add();
            ws.Name = "Subject " + subjectNumber;

            int row = 1;

            ws.Cells[row, 1] = "Heading";
            ws.Cells[row, 2] = "Pitch";
            ws.Cells[row, 3] = "Roll";
            ws.Cells[row, 4] = "Pos-X";
            ws.Cells[row, 5] = "Pos-Y";
            ws.Cells[row, 6] = "Pos-Z";
            ws.Cells[row, 7] = "Vel-X";
            ws.Cells[row, 8] = "Vel-Y";
            ws.Cells[row, 9] = "Vel-Z";
            ws.Cells[row, 10] = "Time";
            ws.Cells[row, 11] = "Samples";
            ws.Cells[row, 12] = "Stops";

            //variables used for calculations- flags and arrays, as well as mini-globals
            Boolean header = true;
            //HPR vars
            Boolean startHPR = false;
            double[] hpr0 = { 0, 0, 0 };
            double[] hpr1 = { 0, 0, 0 };
            double[] delta = { 0, 0, 0 };
            //XYZ var
            double maxPos = 0;
            //<XYZ> var
            double speedThreshold = 0.05;

            //Mean variable
            //H P R X Y Z X Y Z
            double[] mean = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            //0 1 2 3 4 5 6 7 8

            String time = "";
            int[] offset = { 0, 0 };
            Boolean startOff = false;

            //list to speed up bitmap image drawing
            List<int[]> coords = new List<int[]>();

            //shifting row to give space for final results
            row += 5;
            int size = 2000;
            //row 2 = mean
            //row 3 = stddevp
            //row 4 = std error
            //row 5 = special - HPR absval change total, XYZ total distance traveled, VEL n/a, time spent in environ
            //foreach (string a in scans)
            foreach (string a in sample)
            {
                string[] b = a.Split(' ');
                //strips date header
                if (header)
                {
                    var temp = new List<String>(b);
                    temp.RemoveAt(0);
                    temp.RemoveAt(0);
                    temp.RemoveAt(0);
                    b = temp.ToArray();
                    header = false;
                }
                else
                {
                    //sneaky null char or space or something, this removes it
                    var temp = new List<String>(b);
                    temp.RemoveAt(0);
                    b = temp.ToArray();
                }

                Boolean halt = false;
                //checks if this is a halt scan
                foreach (String h in b) { if (h == "!") { halt = true; } }

                //H P R , X Y Z , XX YY ZZ :  HH MM SS mm 
                //0 1 2 3 4 5 6 7 8  9  10 11 12 13 14 15
                if (!halt)
                {
                    //clean up each substring
                    for (int d = 0; d < 17; d++) { b[d] = b[d].Trim(); }
                    //HPR section-----------------------------------------------------------------------------------------
                    //calculating total change in hpr values
                    hpr0[0] = Convert.ToDouble(b[0]);
                    hpr0[1] = Convert.ToDouble(b[1]);
                    hpr0[2] = Convert.ToDouble(b[2]);

                    if (startHPR)
                    {
                        //uses a 2D array to store old (1) and current (0) HPR
                        //then takes difference to calculate total change
                        delta[0] += Math.Abs(hpr1[0] - hpr0[0]);
                        delta[1] += Math.Abs(hpr1[1] - hpr0[1]);
                        delta[2] += Math.Abs(hpr1[2] - hpr0[2]);
                        hpr1 = hpr0;
                    }
                    else if (!startHPR) { hpr1 = hpr0; startHPR = true; }//initial values
                    //file write
                    ws.Cells[row, 1] = b[0];
                    ws.Cells[row, 2] = b[1];
                    ws.Cells[row, 3] = b[2];
                    //add for mean
                    mean[0] += Convert.ToDouble(b[0]);
                    mean[1] += Convert.ToDouble(b[1]);
                    mean[2] += Convert.ToDouble(b[2]);
                    //----------------------------------------------------------------------------------------------------
                    //XYZ position section--------------------------------------------------------------------------------
                    //easy, just record largest value for bmp writing
                    if (Math.Abs(Convert.ToDouble(b[4])) > maxPos)
                    {
                        maxPos = Math.Abs(Convert.ToDouble(b[4]));
                    }
                    else if (Math.Abs(Convert.ToDouble(b[5])) > maxPos)
                    {
                        maxPos = Math.Abs(Convert.ToDouble(b[5]));
                    }
                    else if (Math.Abs(Convert.ToDouble(b[6])) > maxPos)
                    {
                        maxPos = Math.Abs(Convert.ToDouble(b[6]));
                    }
                    //file write
                    ws.Cells[row, 4] = b[4];
                    ws.Cells[row, 5] = b[5];
                    ws.Cells[row, 6] = b[6];
                    //load into list for path image
                    coords.Add(new int[] { (int)Convert.ToDouble(b[4]), (int)Convert.ToDouble(b[5]) });

                    //add to mean
                    mean[3] += (long)Convert.ToDouble(b[4]);
                    mean[4] += (long)Convert.ToDouble(b[5]);
                    mean[5] += (long)Convert.ToDouble(b[6]);
                    //----------------------------------------------------------------------------------------------------
                    //<XYZ> vector section--------------------------------------------------------------------------------
                    if (Convert.ToDouble(b[8]) < speedThreshold && Convert.ToDouble(b[9]) < speedThreshold && Convert.ToDouble(b[10]) < speedThreshold)
                    {
                        //if all vectors are lower than threshold, flag
                        ws.Cells[row, 12] = "1";
                    }
                    else
                    {
                        ws.Cells[row, 12] = "0";
                    }
                    //file write
                    ws.Cells[row, 7] = b[8];
                    ws.Cells[row, 8] = b[9];
                    ws.Cells[row, 9] = b[10];
                    //add to mean
                    mean[6] += (long)Convert.ToDouble(b[8]);
                    mean[7] += (long)Convert.ToDouble(b[9]);
                    mean[8] += (long)Convert.ToDouble(b[10]);
                    //----------------------------------------------------------------------------------------------------
                    //if we're in the time section------------------------------------------------------------------------
                    time = b[12] + ":" + b[13] + ":" + b[14];
                    ws.Cells[row, 10] = time;
                    time = "";
                    //----------------------------------------------------------------------------------------------------
                    //if we're in the counter/millisecond section
                    //building the counter
                    if (!startOff)
                    {
                        //offset[0] is the current count
                        //offset[1] holds the last millisecond read
                        offset[0] = 1;
                        offset[1] = Convert.ToInt32(b[15]);
                        ws.Cells[row, 11] = offset[0];
                        startOff = true;
                    }
                    else
                    {
                        if (Convert.ToInt32(b[15]) != offset[1])
                        {
                            offset[0]++;
                        }
                        //write & update offset[1]
                        ws.Cells[row, 11] = offset[0];
                        offset[1] = Convert.ToInt32(b[15]);
                    }
                    row++;
                }
            }
            String bottom = Convert.ToString(row - 1);
            int scans = row - 6;
            ws.Cells[5, 1] = delta[0];//total change in look vectors
            ws.Cells[5, 2] = delta[1];
            ws.Cells[5, 3] = delta[2];
            //means
            ws.Cells[2, 1] = mean[0] / scans;
            ws.Cells[2, 2] = mean[1] / scans;
            ws.Cells[2, 3] = mean[2] / scans;
            ws.Cells[2, 4] = mean[3] / scans;
            ws.Cells[2, 5] = mean[4] / scans;
            ws.Cells[2, 6] = mean[5] / scans;
            ws.Cells[2, 7] = mean[6] / scans;
            ws.Cells[2, 8] = mean[7] / scans;
            ws.Cells[2, 9] = mean[8] / scans;
            //total time for test
            ws.Cells[2, 10] = "=TEXT(J" + bottom + "-J6, \"hh:mm:ss\")";
            //total number of samples
            ws.Cells[2, 11] = bottom;
            // calculate number of stops, total stopped time, longest stop |
            //load in the sample count and the stopped/not stopped list
            Range stopList = ws.get_Range("L6", ("L" + bottom));
            Range timeList = ws.get_Range("K6", ("K" + bottom));
            Boolean inStop = false;
            int numStops = 0;
            int stopTime = 0;
            int thisStop = 0;
            int longStop = 0;
            //for each set of moving and sample count
            for (int counter = 0; counter < (row - 6); counter++)
            {
                //if stopped
                if (stopList[counter].Value2 == 1)
                {
                    //and previously moving
                    if (inStop == false)
                    {
                        //start counting
                        inStop = true;
                        numStops++;
                        thisStop++;
                    }
                    //else if continuing a stop
                    else if (inStop == true)
                    {
                        //if the time list says we have progressed at least one millisecond increment else do nothing
                        if (timeList[counter - 1].Value2 != timeList[counter].Value2)
                        {
                            thisStop++;
                        }
                    }
                }
                //if we're moving
                else if (stopList[counter].Value2 == 0)
                {
                    //and previously we were stopped
                    if (inStop == true)
                    {
                        //increase stop time total, check if this stop was the longest, zero out thisStop
                        stopTime += thisStop;
                        if (thisStop > longStop) { longStop = thisStop; }
                        thisStop = 0;
                        inStop = false;
                    }
                    //else do nothing
                }
            }
            //formatting and displaying number of stops, total stopped time in seconds, and longest stop in milliseconds
            ws.Cells[2, 12] = numStops;
            ws.Cells[3, 13] = "Time (s)";
            ws.Cells[3, 12] = stopTime / 100.0;
            ws.Cells[4, 13] = "Long (s)";
            ws.Cells[4, 12] = longStop / 100.0;
            //standard deviation and error calculation
            for (int z = 0; z < 9; z++)
            {
                ws.Cells[2, (z + 1)] = mean[z] / (long)(row - 6);
                char col2 = (char)(65 + z);//generates ascii letters A-H without having to hardcode each one
                Range root = (Range)ws.Cells[row - 1, 11];
                ws.Cells[3, (z + 1)] = "=STDEVP(" + col2 + "6:" + col2 + bottom + ")";
                ws.Cells[4, (z + 1)] = "=STDEVP(" + col2 + "6:" + col2 + bottom + ") / SQRT(" + Convert.ToString(root.Value2 + 1) + ")";
            }
            //S x S bmp, remapped to +/-1000 ranges and offset later on to make 0,0 the center of the image which handles negative coordiates
            Bitmap bm = new Bitmap(1, 1);
            bm.SetPixel(0, 0, Color.Black);
            Bitmap bmp = new Bitmap(bm, size, size);

            //rescale the input (x,y) to fit size so that max value position in the file maps to 2000
            double scaling = size / maxPos;
            //reduce it by half to refit it to max out at 1000
            scaling /= 2;
            //for each XY location, set the equivalent pixel to white
            foreach (int[] pix in coords)
            {
                bmp.SetPixel((int)Math.Floor(pix[0] * scaling) + (size / 2), (int)Math.Floor(pix[1] * scaling) + (size / 2), Color.White);
            }
            //save image
            bmp.Save("subject" + Convert.ToString(subjectNumber) + "path.bmp");
            //return object
            return Task.FromResult(bmp);
        }
        /// <summary>
        /// Called once after all processor stuff has finished, combines images into aggregate and then displays excel
        /// </summary>
        /// <param name="num"></param>
        public void FinalizeAggregate(int num)
        {
            //remove blank default sheet
            workbook.Sheets[num].Delete();
            //Initialize new bmp
            Bitmap average = new Bitmap(1, 1);
            //set it to solid black
            average.SetPixel(0, 0, Color.Black);
            Bitmap averagePath = new Bitmap(average, 2000, 2000);

            //Averaging pixel values
            foreach (Bitmap b in bmpAggregate)
            {
                for (int xx = 0; xx < b.Height; xx++)
                {
                    for (int yy = 0; yy < b.Width; yy++)
                    {
                        //default density will increase by an even amount based on number of users
                        //ex 5 users means a point that 5 people crossed should be solid white with a total
                        //of 5 levels of brightness
                        int density = (int)Math.Floor((255 / bmpAggregate.Capacity) + 0.5);
                        if ((density * (bmpAggregate.Capacity - 1)) <= 255)
                        {
                            //this makes sure any point all paths cross over will be totally white, not light grey
                            //aka accounting for rounding errors by double checking and then jacking it up
                            density += 20;
                        }
                        Color oldCol = averagePath.GetPixel(xx, yy);
                        Color pixCol;
                        if (b.GetPixel(xx, yy).R > 0 && b.GetPixel(xx, yy).G > 0 && b.GetPixel(xx, yy).B > 0)
                        {
                            int red = oldCol.R + density;
                            int green = oldCol.G + density;
                            int blue = oldCol.B + density;
                            if (red > 255) { red = 255; }
                            if (green > 255) { green = 255; }
                            if (blue > 255) { blue = 255; }
                            pixCol = Color.FromArgb(red, green, blue);
                        }
                        else
                        {
                            pixCol = oldCol;
                        }
                        averagePath.SetPixel(xx, yy, pixCol);
                    }
                }
            }
            averagePath.Save("heatmap.bmp");
            //displays excel sheet
            excel.Visible = true;
        }
    }
}
