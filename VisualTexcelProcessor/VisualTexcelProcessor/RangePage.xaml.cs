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
    /// Interaction logic for RangePage.xaml
    /// </summary>
    public partial class RangePage : Page
    {
        public RangePage()
        {
            InitializeComponent();
        }
        //every time something is typed, check that it's valid
        private void MinBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!CheckIfNumbers(Convert.ToChar(e.Text)))
            {
                MessageBox.Show("Error: Lower file number must be a number");
                MinBox.Text = "";
            }
        }
        //every time something is typed, check that it's valid
        private void MaxBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!CheckIfNumbers(Convert.ToChar(e.Text)))
            {
                MessageBox.Show("Error: Higher file number must be a number");
                MaxBox.Text = "";
            }
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            //extra verification of inputs
            if (!(CheckIfNumbers(MinBox.Text) && CheckIfNumbers(MaxBox.Text)))
            {
                MessageBox.Show("Invalid file numbers");
                return;
            }
            int max = int.Parse(MaxBox.Text);
            int min = int.Parse(MinBox.Text);
            //handle it if wrong inputs are given
            if (min > max)
            {
                int temp = min;
                min = max;
                max = temp;
            }
            List<bool> subjects = new List<bool>();
            Processor machine = new Processor();
            for (int x = min; x < max; x++)
            {
                MessageBox.Show("Starting subject " + x);
                bool subjectBool = await machine.FileProcessor(x);
                subjects.Add(subjectBool);
                MessageBox.Show("Finished subject " + x);
            }
            MessageBox.Show("Outside of await loop");
            //machine.FinalizeAggregate(max);
        }
        //overloaded to accept a char or a string
        private bool CheckIfNumbers(string input)
        {
            int val;
            if (int.TryParse(input, out val))
                return true;
            else
                return false;
        }

        private bool CheckIfNumbers(Char e)
        {
            if (Char.IsNumber(e))
                return true;
            else
                return false;
        }
    }
}
