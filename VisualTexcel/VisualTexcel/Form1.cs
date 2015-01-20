using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace VisualTexcel
{
    public partial class GUI : Form
    {
        // Startup and run --------------------------------------------------------
        // Declare variables for processing -----------------
        //Excel object- contains workbooks
        private static Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        //Workbook object- contains sheets
        private static Workbook wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        //Aggregate BMP image, has generated BMPs added to it which are later overlaid to make a heatmap
        List<Bitmap> bmpAggregate = new List<Bitmap>();
        //---------------------------------------------------
        // Declare variables for GUI-------------------------
        // File selection
        int min = 0;
        int max = 0;
        //this is used to delete blank default sheet
        int num = 2;
        string listFiles = "";
        //--------------------------------------------------
        public GUI()
        {
            //check that excel and workbook loaded
            if (excel == null)
            {
                MessageBox.Show("Excel could not be started.");
                return;
            }
            else if (wb == null)
            {
                MessageBox.Show("Workbook could not be created.");
                return;
            }
            //if we make it here, it's safe to start the program
            else
            {
                InitializeComponent();
            }
        }
        private void GUI_Load(object sender, EventArgs e)
        {
            //On startup
            //set progress bar range
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
        }
        // Start buttons ----------------------------------------------------------
        private void rangeStart_Click(object sender, EventArgs e)
        {
            //remove other buttons
            Clear();
            //lock inputs
            minFileNum.Controls.Clear();
            maxFileNum.Controls.Clear();

            //remove other options
            SingleFileLabel.Dispose();
            singleFileSelectLabel.Dispose();
            singleFileNum.Dispose();
            ListofFilesLabel.Dispose();
            listInstructionLabel.Dispose();
            listString.Dispose();

            //resize
            GUI.ActiveForm.Width = GUI.ActiveForm.Width/2;

            // Begin processing -----------------------------------
            if (min > max)
            {
                int temp = min;
                min = max;
                max = temp;                
            }
            for (int x = max; x >= min; x--)
            {

                progressMessage1.Text = "Subject " + x;
                bmpAggregate.Add(excelWrite(x));
                num++;
            }
            FinalizeAggregate();
            progressMessage1.Text = "Complete";
            //-----------------------------------------------------
        }

        private void listStart_Click(object sender, EventArgs e)
        {
            //remove other buttons
            Clear();
            //lock inputs
            listString.ReadOnly = true;

            //remove other options
            SingleFileLabel.Dispose();
            singleFileSelectLabel.Dispose();
            singleFileNum.Dispose();
            RangeofFilesLabel.Dispose();
            minFileNumLabel.Dispose();
            maxFileNumLabel.Dispose();
            minFileNum.Dispose();
            maxFileNum.Dispose();

            //resize
            GUI.ActiveForm.Height = (GUI.ActiveForm.Height / 5)*3;

            // Begin processing -----------------------------------
            String[] list = listFiles.Split(' ');
            foreach (String n in list)
            {
                num++;
                progressMessage1.Text = "Subject " + (int)Convert.ToDouble(n);
                bmpAggregate.Add(excelWrite((int)Convert.ToDouble(n)));
            }
            FinalizeAggregate();
            progressMessage1.Text = "Complete";
            //-----------------------------------------------------
        }

        private void singleStart_Click(object sender, EventArgs e)
        {
            //remove other buttons
            Clear();
            //lock inputs
            singleFileNum.Controls.Clear();

            //remove other options
            ListofFilesLabel.Dispose();
            listInstructionLabel.Dispose();
            listString.Dispose();
            RangeofFilesLabel.Dispose();
            minFileNumLabel.Dispose();
            maxFileNumLabel.Dispose();
            minFileNum.Dispose();
            maxFileNum.Dispose();

            //resize
            GUI.ActiveForm.Width = GUI.ActiveForm.Width / 2;

            // Begin processing -----------------------------------
            // pretty easy for single files
            progressMessage1.Text = "Subject " + min;
            bmpAggregate.Add(excelWrite(min));
            FinalizeAggregate();
            progressMessage1.Text = "Completed";
            //-----------------------------------------------------
        }

        //-------------------------------------------------------------------------
        // File selectors ---------------------------------------------------------
        private void maxFileNum_ValueChanged(object sender, EventArgs e)
        {
            max = (int)maxFileNum.Value;
        }

        private void minFileNum_ValueChanged(object sender, EventArgs e)
        {
            min = (int)minFileNum.Value;
        }

        private void singleFileNum_ValueChanged(object sender, EventArgs e)
        {
            min = (int)singleFileNum.Value;
        }

        private void listString_Changed(object sender, EventArgs e)
        {
            listFiles = listString.Text;
        }
//******************************************************************************************************************
        //this does the actual work, int row probably isnt needed because of multi-sheet solution
        public Bitmap excelWrite(int subj)
        {
            progressBar.Value = 0;
            string subject = "subject" + subj + ".txt";
            progressMessage2.Text = "Finding " + subject + "...";

            StreamReader sr = new StreamReader(subject);

            //I *think* this should handle missing files without halting the program
            //ex. 1 2 3 5 6 7 on a sequential 1-7 should just skip 4 and keep going
            //**UNTESTED**x
            if (sr != null) { progressMessage2.Text = "File found!"; }
            else { progressMessage2.Text = "Error: " + subject + " not found!"; return null; }
            
            string raw = sr.ReadToEnd();
            string[] sample = raw.Split(';');

            Worksheet ws = wb.Worksheets.Add();
            ws.Name = "Subject " + subj;
            progressMessage2.Text = "Making worksheet "+subj+"...";
            if (ws == null)
            {
                progressMessage2.Text = "Worksheet could not be created.";
            }

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
            double[] hpr0 = {0,0,0};
            double[] hpr1 = {0,0,0};
            double[] delta = { 0, 0, 0 };
            //XYZ var
            double maxPos = 0;
            //<XYZ> var
            double speedThreshold = 0.05;

            //Mean variable
                           //H P R X Y Z X Y Z
            double[] mean = {0,0,0,0,0,0,0,0,0};
                           //0 1 2 3 4 5 6 7 8
            
            String time = "";
            int[] offset = {0, 0};
            Boolean startOff = false;

            //list to speed up bitmap image drawing
            List<int[]> coords = new List<int[]>();

            //shifting row to give space for final math-ed stuff
            row += 5;
            progressMessage2.Text = "Writing File";
            int size = 2000;
            int sus = 0;
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

                sus++;
                if (sus % 100 == 0)
                {
                    if (progressBar.Value < 100)
                    {
                        progressBar.Value++;
                    }
                    else
                    {
                        progressBar.Value = 0;
                    }
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
            progressBar.Value = 0;
            progressMessage2.Text = "Secondary File Processing";
            String bottom = Convert.ToString(row-1);
            int scans = row - 6;
            ws.Cells[5, 1] = delta[0];//total change in look vectors
            ws.Cells[5, 2] = delta[1];
            ws.Cells[5, 3] = delta[2];
            //means
            progressBar.Value += 10;
            ws.Cells[2, 1] = mean[0] / scans;
            ws.Cells[2, 2] = mean[1] / scans;
            ws.Cells[2, 3] = mean[2] / scans;
            ws.Cells[2, 4] = mean[3] / scans;
            ws.Cells[2, 5] = mean[4] / scans;
            ws.Cells[2, 6] = mean[5] / scans;
            ws.Cells[2, 7] = mean[6] / scans;
            ws.Cells[2, 8] = mean[7] / scans;
            ws.Cells[2, 9] = mean[8] / scans;
            progressBar.Value += 10;
            //total time for test
            ws.Cells[2, 10] = "=TEXT(J" + bottom + "-J6, \"hh:mm:ss\")";
            //total number of samples
            ws.Cells[2, 11] = bottom;
            progressBar.Value += 10;
            // calculate number of stops, total stopped time, longest stop |
            //load in the sample count and the stopped/not stopped list
            Range stopList = ws.get_Range("L6", ("L" + bottom));
            Range timeList = ws.get_Range("K6", ("K" + bottom));
            Boolean inStop = false;
            int numStops = 0;
            int stopTime = 0;
            int thisStop = 0;
            int longStop = 0;
            progressBar.Value += 10;
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
            progressBar.Value += 20;
            //formatting and displaying number of stops, total stopped time in seconds, and longest stop in milliseconds
            ws.Cells[2, 12] = numStops;
            ws.Cells[3, 13] = "Time (s)";
            ws.Cells[3, 12] = stopTime / 100.0;
            ws.Cells[4, 13] = "Long (s)";
            ws.Cells[4, 12] = longStop / 100.0;
            progressBar.Value += 20;
            //standard deviation and error calculation
            for (int z = 0; z < 9;z++)
            {
                ws.Cells[2, (z+1)] = mean[z] / (long) (row-6);
                char col2 = (char) (65 + z);//generates ascii letters A-H without having to hardcode each one
                Range root = (Range)ws.Cells[row-1, 11];
                ws.Cells[3, (z + 1)] = "=STDEVP(" + col2 + "6:" + col2 + bottom + ")";
                ws.Cells[4, (z + 1)] = "=STDEVP(" + col2 + "6:" + col2 + bottom + ") / SQRT("+Convert.ToString(root.Value2+1)+")";
            }
            progressBar.Value += 20;
            //notifications!
            progressMessage2.Text = "Generating Path Image";
            progressBar.Value = 0;
            //S x S bmp, remapped to +/-1000 ranges and offset later on to make 0,0 the center of the image which handles negative coordiates
            Bitmap bm = new Bitmap(1, 1);
            bm.SetPixel(0, 0, Color.Black);
            Bitmap bmp = new Bitmap(bm, size, size);

            progressMessage2.Text = "Drawing Path";
            progressBar.Value = 0;
            //rescale the input (x,y) to fit size so that max value position in the file maps to 2000
            double scaling = size / maxPos;
            //reduce it by half to refit it to max out at 1000
            scaling /= 2;
            //for each XY location, set the equivalent pixel to white
            int p = 0;
            foreach(int[] pix in coords)
            {
                if (p % (coords.Capacity / 10) == 0) 
                {
                    if (progressBar.Value < 100)
                    {
                        progressBar.Value++;
                    }
                    else
                    {
                        progressBar.Value = 0;
                    }
                }
                p++;
                bmp.SetPixel((int)Math.Floor(pix[0]*scaling)+(size/2), (int)Math.Floor(pix[1]*scaling)+(size/2), Color.White);
            }
            progressBar.Value = 100;
            //save image
            bmp.Save("subject"+Convert.ToString(subj)+"path.bmp");
            //return object
            return bmp;
        }
//******************************************************************************************************************
        // Aggregate Image and finishing-------------------------------------------
        private void FinalizeAggregate()
        {
            //remove blank default sheet
            wb.Sheets[num].Delete();
            //Initialize new bmp
            progressBar.Value = 0;
            progressMessage2.Text = "Aggregate Image Initializing";
            Bitmap average = new Bitmap(1,1);
            //set it to solid black
            average.SetPixel(0, 0, Color.Black);
            Bitmap averagePath = new Bitmap(average, 2000, 2000);

            //Averaging pixel values
            int bit = 1;
            foreach (Bitmap b in bmpAggregate)
            {
                progressMessage2.Text = "Generating Heat Map "+bit;
                progressBar.Value = 0;
                for (int xx = 0; xx < b.Height; xx++)
                {
                    if (xx % 200 == 0) { progressBar.Value += 10; }
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
                bit++;
            }
            progressMessage2.Text = "Heat Map Complete";
            averagePath.Save("heatmap.bmp");
            //displays excel sheet
            excel.Visible = true;
        }
        //-------------------------------------------------------------------------
        // Cleanup function -------------------------------------------------------
        // this just removes the buttons to prevent more than one mode being selected
        private void Clear()
        {
            rangeStart.Dispose();
            singleStart.Dispose();
            listStart.Dispose();
        }
        //-------------------------------------------------------------------------
        // Info functions - text wall incoming ------------------------------------
        private void menu_help_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Help\n\n"+
                "Troubleshooting\n"+
                "Did it seem to freeze or lock out while processing? That's normal.\n"+
                "This program requires that Microsoft Office Excel is installed to function.\n"+
                "Make sure that VisualTexcel.exe is in the same folder as the subject files.\n"+
                "Subject files must follow the following format exactly:\n"+
                "DD-MM-YY File Begin: H P R , X Y Z , I J K : HH MM SS mm ; ! Halt: Scans Q\n"+
                "Where everything between 'Begin:' and '!' is repeated Q times, and the data is:\n"+
                "Heading, Pitch, Roll, Location(X,Y,Z), Velocity(I,J,K), Hr Min Sec Millisecond\n"+
                "And the files must be named 'SubjectX.txt' where X is the subject number.\n\n"+
                "Instructions\n" +
                "VisualTexcel has three modes: Range, List, and Single.\n"+
                "Range processes every file numbered from Start file to End file, inclusive.\n"+
                "List processes each file number specified- this is for out of order files, \n"+
                "ex. 1 3 6 9 10 11\n"+
                "Single processes one file, specified by the file number.\n"+
                "Only one mode can be run at a time. Once a Process button has been pressed,\n"+
                "data in the settings for any other mode is ignored. For example, entering\n"+
                "numbers into every mode settings and then pressing 'Process Range' will\n"+
                "only use the data in 'File to start at' and 'File to end at'. Furthermore,\n"+
                "once started, the selected files can not be changed.\n\n"+
                "Description of Output\n" +
                "The output is an excel spreadsheet listing the mean, standard deviation,\n"+
                "and standard error for each section. In addition, it lists the delta for\n"+
                "the heading, pitch, and roll, giving the total amount of change.\n"+
                "Finally, it also lists the number of times the subject stopped moving,\n"+
                "as well as the total time spent standing still and the longest pause.\n"+
                "Each subject in a processing batch is put into the same workbook, under\n"+
                "different worksheets. Worksheets will be sequential in the order they were\n"+
                "entered in the case of List, or just sequential in the case of Range.\n"+
                "The program also outputs BMP Images of the path taken through the test,\n"+
                "one for each subject, as well as one aggregate heatmap image of all the\n"+
                "paths overlaid. These images will be placed in the same location as the\n"+
                "subject files.");
        }

        private void menu_about_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Code written by Tyler Harbin-Giuntoli\n"+
                "This program is free to use for research purposes (it's not good for anything else)\n"+
                "as long as attribution is given in any research produced (stick in a line saying\n"+
                "'Processed using VisualTexcel by Tyler Harbin-Giuntoli' or something)\n"+
                "Originally written for the College of Sciences at the University of Central Florida\n"+
                "This was written in C# using Microsoft Visual Studio 2013, source code is available\n"+
                "upon request by emailing Harbin.Giuntoli@knights.ucf.edu\n"+
                "or Harbin.Giuntoli@gmail.com. That's also probably a good place\n"+
                "to send any bug reports or if you want me to write a special version\n"+
                "for you/your research. For any emails, please use subject 'VisualTexcel'.");
        }
        //-------------------------------------------------------------------------
    }
}
