using System;
using System.Collections.Generic;
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

namespace VisualTexcelProcessor
{
    /// <summary>
    /// Interaction logic for StartPage.xaml
    /// </summary>
    public partial class StartPage : Page
    {
        public StartPage()
        {
            InitializeComponent();
        }

        private void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            About about = new About();
            about.Show();
        }

        private void SingleButton_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Navigate(new SinglePage());
        }

        private void RangeButton_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Navigate(new RangePage());
        }

        private void ListButton_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Navigate(new RangePage());
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
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
    }
}
