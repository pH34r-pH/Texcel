using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace TexcelApplication
{
    /// <summary>
    /// Interaction logic for StartPage.xaml
    /// </summary>
    public partial class StartPage : System.Windows.Controls.Page
    {
        public StartPage()
        {
            InitializeComponent();
            excel = new Microsoft.Office.Interop.Excel.Application();
            workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        }
        //---------------------------------------------------------------


        // Startup and run --------------------------------------------------------
        // Declare variables for processing -----------------
        //Excel object- contains workbooks
        private static Microsoft.Office.Interop.Excel.Application excel;
        //Workbook object- contains sheets
        private static Workbook workbook;
        //Aggregate BMP image, has generated BMPs added to it which are later overlaid to make a heatmap
        List<Bitmap> bmpAggregate = new List<Bitmap>();
        //---------------------------------------------------
        // Declare variables for GUI-------------------------
        // File selection
        int num = 0;
        int min = 0;
        int max = 0;
        //--------------------------------------------------
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (excel == null)
            {
                MessageBox.Show("Excel could not be started.");
                return;
            }
            else if (workbook == null)
            {
                MessageBox.Show("Workbook could not be created.");
                return;
            }
            Progress.Value = 0;
            ProcessUITasks();
            if (RangeButton.IsChecked.Value == true)
            {
                try
                {
                    min = (int)Convert.ToDouble(MinBox.Text);
                    max = (int)Convert.ToDouble(MaxBox.Text);
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Error: " + exp.Message);
                    return;
                }
                if (min > max)
                {
                    int temp = min;
                    min = max;
                    max = temp;
                }
                for (int x = max; x >= min; x--)
                {
                    bmpAggregate.Add(ExcelWrite(x));
                }
                FinalizeAggregate();
            }
            else if (ListButton.IsChecked.Value == true)
            {
                string[] list;
                List<string> SubjectList = new List<string>();
                try
                {
                    list = ListBox.Text.Split(' ');
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Error: " + exp.Message);
                    return;
                }
                
                foreach (string n in list)
                {
                    if (n.Length > 0)
                    {
                        SubjectList.Add(n);
                        MessageBox.Show("added " + n);
                    }
                }
                foreach (string n in SubjectList)
                {
                    bmpAggregate.Add(ExcelWrite((int)Convert.ToDouble(n)));
                }
                FinalizeAggregate();
            }
            else if (SingleButton.IsChecked.Value == true)
            {
                try
                {
                    min = (int)Convert.ToDouble(SingleBox.Text);
                }
                catch(Exception exp)
                {
                    MessageBox.Show("Error: " + exp.Message);
                    return;
                }
                bmpAggregate.Add(ExcelWrite(min));
                FinalizeAggregate();
            }

        }

//******************************************************************************************************************
        public Bitmap ExcelWrite(int SubjectNumber)
        {
            Progress.Value = 0;
            ProcessUITasks();
            var SubjectFile = "subject" + SubjectNumber + ".txt";
            StreamReader sr;
            if (File.Exists(SubjectFile))
            {
                sr = new StreamReader(SubjectFile);
            }
            else
            {
                MessageBox.Show("Subject " + SubjectNumber + " not found; skipping.");
                num--;
                return null;
            }

            string RawString = sr.ReadToEnd();
            string[] SampleArray = RawString.Split(';');

            Worksheet Sheet = workbook.Worksheets.Add();
            Sheet.Name = "Subject " + SubjectNumber;

            int row = 1;

            Sheet.Cells[row, 1] = "Heading";
            Sheet.Cells[row, 2] = "Pitch";
            Sheet.Cells[row, 3] = "Roll";
            Sheet.Cells[row, 4] = "Pos-X";
            Sheet.Cells[row, 5] = "Pos-Y";
            Sheet.Cells[row, 6] = "Pos-Z";
            Sheet.Cells[row, 7] = "Vel-X";
            Sheet.Cells[row, 8] = "Vel-Y";
            Sheet.Cells[row, 9] = "Vel-Z";
            Sheet.Cells[row, 10] = "Time";
            Sheet.Cells[row, 11] = "Samples";
            Sheet.Cells[row, 12] = "Stops";

            // Variables used for calculations- flags and arrays, as well as mini-globals
            Boolean Header = true;
            // HPR vars
            Boolean StartedHPR = false;
            double[] hprThis = { 0, 0, 0 };
            double[] hprLast = { 0, 0, 0 };
            double[] delta = { 0, 0, 0 };
            // XYZ var
            double MaxPos = 0;
            // <XYZ> var
            double SpeedThreshold = 0.05;

            //Mean variable
            //H P R X Y Z X Y Z
            double[] mean = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            //0 1 2 3 4 5 6 7 8

            String time = "";
            int[] offset = { 0, 0 };
            Boolean StartIsOff = false;

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
            int Increment = 100 / SampleArray.Length;
            foreach (string a in SampleArray)
            {
                Progress.Value += Increment;
                ProcessUITasks();
                string[] b = a.Split(' ');
                //strips date header
                if (Header)
                {
                    var temp = new List<String>(b);
                    temp.RemoveAt(0);
                    temp.RemoveAt(0);
                    temp.RemoveAt(0);
                    b = temp.ToArray();
                    Header = false;
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
                    //HPR section----------------------------------------------------------------------------------------
                    //calculating total change in hpr values
                    hprThis[0] = Convert.ToDouble(b[0]);
                    hprThis[1] = Convert.ToDouble(b[1]);
                    hprThis[2] = Convert.ToDouble(b[2]);

                    if (StartedHPR)
                    {
                        //uses a 2D array to store old (1) and current (0) HPR
                        //then takes difference to calculate total change

                        System.Windows.MessageBox.Show("Delta: " + delta[0] + " abs value: " + Math.Abs((hprThis[0] - hprLast[0])) + " HPR values: " +hprThis[0] + " " + hprLast[0]);
                        delta[0] += Math.Abs(hprLast[0] - hprThis[0]);
                        delta[1] += Math.Abs(hprLast[1] - hprThis[1]);
                        delta[2] += Math.Abs(hprLast[2] - hprThis[2]);
                    }
                    // This lets us assign initial values.
                    else if (!StartedHPR) 
                    {      
                        StartedHPR = true; 
                    }
                    Array.Copy(hprThis, hprLast, hprLast.Length);

                    //file write
                    Sheet.Cells[row, 1] = b[0];
                    Sheet.Cells[row, 2] = b[1];
                    Sheet.Cells[row, 3] = b[2];
                    //add for mean
                    mean[0] += Convert.ToDouble(b[0]);
                    mean[1] += Convert.ToDouble(b[1]);
                    mean[2] += Convert.ToDouble(b[2]);
                    //---------------------------------------------------------------------------------------------------
                    //XYZ position section-------------------------------------------------------------------------------
                    //easy, just record largest value for bmp writing
                    if (Math.Abs(Convert.ToDouble(b[4])) > MaxPos)
                    {
                        MaxPos = Math.Abs(Convert.ToDouble(b[4]));
                    }
                    else if (Math.Abs(Convert.ToDouble(b[5])) > MaxPos)
                    {
                        MaxPos = Math.Abs(Convert.ToDouble(b[5]));
                    }
                    else if (Math.Abs(Convert.ToDouble(b[6])) > MaxPos)
                    {
                        MaxPos = Math.Abs(Convert.ToDouble(b[6]));
                    }
                    //file write
                    Sheet.Cells[row, 4] = b[4];
                    Sheet.Cells[row, 5] = b[5];
                    Sheet.Cells[row, 6] = b[6];
                    //load into list for path image
                    coords.Add(new int[] { (int)Convert.ToDouble(b[4]), (int)Convert.ToDouble(b[5]) });

                    //add to mean
                    mean[3] += (long)Convert.ToDouble(b[4]);
                    mean[4] += (long)Convert.ToDouble(b[5]);
                    mean[5] += (long)Convert.ToDouble(b[6]);
                    //---------------------------------------------------------------------------------------------------
                    //<XYZ> vector section-------------------------------------------------------------------------------
                    if (Convert.ToDouble(b[8]) < SpeedThreshold && Convert.ToDouble(b[9]) < SpeedThreshold && Convert.ToDouble(b[10]) < SpeedThreshold)
                    {
                        //if all vectors are lower than threshold, flag
                        Sheet.Cells[row, 12] = "1";
                    }
                    else
                    {
                        Sheet.Cells[row, 12] = "0";
                    }
                    //file write
                    Sheet.Cells[row, 7] = b[8];
                    Sheet.Cells[row, 8] = b[9];
                    Sheet.Cells[row, 9] = b[10];
                    //add to mean
                    mean[6] += (long)Convert.ToDouble(b[8]);
                    mean[7] += (long)Convert.ToDouble(b[9]);
                    mean[8] += (long)Convert.ToDouble(b[10]);
                    //---------------------------------------------------------------------------------------------------
                    //if we're in the time section-----------------------------------------------------------------------
                    time = b[12] + ":" + b[13] + ":" + b[14];
                    Sheet.Cells[row, 10] = time;
                    time = "";
                    //---------------------------------------------------------------------------------------------------
                    //if we're in the counter/millisecond section
                    //building the counter
                    if (!StartIsOff)
                    {
                        //offset[0] is the current count
                        //offset[1] holds the last millisecond read
                        offset[0] = 1;
                        offset[1] = Convert.ToInt32(b[15]);
                        Sheet.Cells[row, 11] = offset[0];
                        StartIsOff = true;
                    }
                    else
                    {
                        if (Convert.ToInt32(b[15]) != offset[1])
                        {
                            offset[0]++;
                        }
                        //write & update offset[1]
                        Sheet.Cells[row, 11] = offset[0];
                        offset[1] = Convert.ToInt32(b[15]);
                    }
                    row++;
                }
            }
            String bottom = Convert.ToString(row - 1);
            int scans = row - 6;
            Sheet.Cells[5, 1] = delta[0];//total change in look vectors
            Sheet.Cells[5, 2] = delta[1];
            Sheet.Cells[5, 3] = delta[2];
            //means
            Sheet.Cells[2, 1] = mean[0] / scans;
            Sheet.Cells[2, 2] = mean[1] / scans;
            Sheet.Cells[2, 3] = mean[2] / scans;
            Sheet.Cells[2, 4] = mean[3] / scans;
            Sheet.Cells[2, 5] = mean[4] / scans;
            Sheet.Cells[2, 6] = mean[5] / scans;
            Sheet.Cells[2, 7] = mean[6] / scans;
            Sheet.Cells[2, 8] = mean[7] / scans;
            Sheet.Cells[2, 9] = mean[8] / scans;
            //total time for test
            Sheet.Cells[2, 10] = "=TEXT(J" + bottom + "-J6, \"hh:mm:ss\")";
            //total number of samples
            Sheet.Cells[2, 11] = bottom;
            // calculate number of stops, total stopped time, longest stop |
            //load in the sample count and the stopped/not stopped list
            Range stopList = Sheet.get_Range("L6", ("L" + bottom));
            Range timeList = Sheet.get_Range("K6", ("K" + bottom));
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
            Sheet.Cells[2, 12] = numStops;
            Sheet.Cells[3, 13] = "Time (s)";
            Sheet.Cells[3, 12] = stopTime / 100.0;
            Sheet.Cells[4, 13] = "Long (s)";
            Sheet.Cells[4, 12] = longStop / 100.0;
            //standard deviation and error calculation
            for (int z = 0; z < 9; z++)
            {
                Sheet.Cells[2, (z + 1)] = mean[z] / (long)(row - 6);
                char col2 = (char)(65 + z);//generates ascii letters A-H without having to hardcode each one
                Range root = (Range)Sheet.Cells[row - 1, 11];
                Sheet.Cells[3, (z + 1)] = "=STDEVP(" + col2 + "6:" + col2 + bottom + ")";
                Sheet.Cells[4, (z + 1)] = "=STDEVP(" + col2 + "6:" + col2 + bottom + ") / SQRT(" + Convert.ToString(root.Value2 + 1) + ")";
            }
            //S x S bmp, remapped to +/-1000 ranges and offset later on to make 0,0 the center of the image which handles negative coordiates
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.FillRectangle(new SolidBrush(System.Drawing.Color.Black), 0, 0, bmp.Width, bmp.Height);
            }

            //rescale the input (x,y) to fit size so that max value position in the file maps to 2000
            double scaling = size / MaxPos;
            //reduce it by half to refit it to max out at 1000
            scaling /= 2;
            //for each XY location, set the equivalent pixel to white
            foreach (int[] pix in coords)
            {
                int XVal = (int)Math.Floor(pix[0] * scaling) + (size / 2);
                int YVal = (int)Math.Floor(pix[1] * scaling) + (size / 2);
                if (XVal > (bmp.Height-1))
                {
                    XVal -= (XVal - (bmp.Height-1));
                }
                if (YVal > (bmp.Width-1))
                {
                    YVal -= (YVal - (bmp.Width-1));
                }
                bmp.SetPixel(XVal, YVal, System.Drawing.Color.White);
            }
            //save image
            bmp.Save("subject" + Convert.ToString(SubjectNumber) + "path.bmp");
            //return object
            return bmp;
        }

        /// <summary>
        /// Called once after all processor stuff has finished, combines images into aggregate and then displays excel
        /// </summary>
        public void FinalizeAggregate()
        {
            Progress.Value = 0;
            ProcessUITasks();
            num += bmpAggregate.Count + 1;
            int Increment = 100 / num;
            //remove blank default sheet
            workbook.Sheets[num].Name = "Notes";
            //Initialize new bmp
            Bitmap averagePath = new Bitmap(2000, 2000);
            using (Graphics g = Graphics.FromImage(averagePath))
            {
                g.FillRectangle(new SolidBrush(System.Drawing.Color.Black), 0, 0, averagePath.Width, averagePath.Height);
            }

            //Averaging pixel values
            foreach (Bitmap b in bmpAggregate )
            {
                if (b != null)
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
                            System.Drawing.Color oldCol = averagePath.GetPixel(xx, yy);
                            System.Drawing.Color pixCol;
                            if (b.GetPixel(xx, yy).R > 0 && b.GetPixel(xx, yy).G > 0 && b.GetPixel(xx, yy).B > 0)
                            {
                                int red = oldCol.R + density;
                                int green = oldCol.G + density;
                                int blue = oldCol.B + density;
                                if (red > 255) { red = 255; }
                                if (green > 255) { green = 255; }
                                if (blue > 255) { blue = 255; }
                                pixCol = System.Drawing.Color.FromArgb(red, green, blue);
                            }
                            else
                            {
                                pixCol = oldCol;
                            }
                            averagePath.SetPixel(xx, yy, pixCol);
                        }
                    }
                    Progress.Value += Increment;
                    ProcessUITasks();
                }
            }
            averagePath.Save("heatmap.bmp");
            //displays excel sheet
            excel.Visible = true;
        }


        //---------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// This should force update the UI
        /// </summary>
        public static void ProcessUITasks()
        {
            var frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(delegate(object parameter)
            {
                frame.Continue = false;
                return null;
            }), null);
            Dispatcher.PushFrame(frame);
        }

        /// <summary>
        /// Sets mode to Range, displays range options and hides all others.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RangeMode(object sender, RoutedEventArgs e)
        {
            try
            {
                RangeMaxLabel.Visibility = Visibility.Visible;
                RangeMinLabel.Visibility = Visibility.Visible;
                MinBox.Visibility = Visibility.Visible;
                MaxBox.Visibility = Visibility.Visible;

                ListLabel.Visibility = Visibility.Hidden;
                ListLabel_2.Visibility = Visibility.Hidden;
                ListBox.Visibility = Visibility.Hidden;
                SingleLabel.Visibility = Visibility.Hidden;
                SingleBox.Visibility = Visibility.Hidden;
            }
            catch(Exception)
            {
                // Do nothing- this catches null exceptions caused by initial assignments
            }
        }

        /// <summary>
        /// Sets mode to list, displays list options and hides all others.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListMode(object sender, RoutedEventArgs e)
        {
            try
            {
                ListLabel.Visibility = Visibility.Visible;
                ListLabel_2.Visibility = Visibility.Visible;
                ListBox.Visibility = Visibility.Visible;

                RangeMaxLabel.Visibility = Visibility.Hidden;
                RangeMinLabel.Visibility = Visibility.Hidden;
                MinBox.Visibility = Visibility.Hidden;
                MaxBox.Visibility = Visibility.Hidden;
                SingleLabel.Visibility = Visibility.Hidden;
                SingleBox.Visibility = Visibility.Hidden;
            }
            catch(Exception)
            {
                // Do nothing- this catches null exceptions
            }
        }

        /// <summary>
        /// Sets mode to single, displays single options and hides all others. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SingleMode(object sender, RoutedEventArgs e)
        {
            try
            {
                SingleLabel.Visibility = Visibility.Visible;
                SingleBox.Visibility = Visibility.Visible;

                RangeMaxLabel.Visibility = Visibility.Hidden;
                RangeMinLabel.Visibility = Visibility.Hidden;
                MinBox.Visibility = Visibility.Hidden;
                MaxBox.Visibility = Visibility.Hidden;
                ListLabel.Visibility = Visibility.Hidden;
                ListLabel_2.Visibility = Visibility.Hidden;
                ListBox.Visibility = Visibility.Hidden;
            }
            catch(Exception)
            {
                // Do nothing- this is just to catch null exceptions
            }
        }

        private void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            AboutBox About = new TexcelApplication.AboutBox();
            About.Show();
        }
    }
}
