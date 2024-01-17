using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DaysheetGenerator_2023
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        
        private void Create_Sheets(object sender, RoutedEventArgs e) //Checks the validity of the files selected and runs the program
        {
            
            if (filePaths.Items.Count <= 0) //Checks if files selected is empty
            {
                string caption = "Error";
                string message = "No files selected to convert!";
                System.Windows.MessageBox.Show(message, caption);
            }
            else if (filePaths.Items.Count > 6) //Checks if more than the maximum number of files is selected
            {
                string caption = "Error";
                string message = "More than 6 files selected!";
                System.Windows.MessageBox.Show(message, caption);
            }

            else //Prompts user to select and output destination folder and runs the program
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();

                string outputFolderPath;
                DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    outputFolderPath = dialog.SelectedPath;
                    Main generator = new Main();
                    generator.RunProgram(filePaths.Items, outputFolderPath);

                    string caption = "Finished";
                    string message = "File has been written to: " + outputFolderPath;
                    System.Windows.MessageBox.Show(message, caption);
                }



            }
            
        }

        
        private void Select_Files(object sender, RoutedEventArgs e) //Opens a dialogue to select lightning bolt .xls files to be used as input
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Office Files|*.xls";

            List<string> pathStrings = new List<string>();
            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true) 
            { 
                foreach (string file in openFileDialog.FileNames) 
                { 
                    filePaths.Items.Add(file);
                }
            }
        }

        
        private void Clear_Files(object sender, RoutedEventArgs e) //Clears selected files from the program
        {
            filePaths.Items.Clear();
        }

        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e) { }
    }
}
