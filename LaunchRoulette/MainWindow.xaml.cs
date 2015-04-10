using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Excel;
using LaunchRoulette.Common;
using Path = System.Windows.Shapes.Path;

namespace LaunchRoulette
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


        /// <summary>
        /// Selects the the file and puts into the path field for Active field.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Active_SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new System.Windows.Forms.OpenFileDialog();
            var result = fileDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var file = fileDialog.FileName;
                    Active_FilePathTextBox.Text = file;
                    Active_FilePathTextBox.ToolTip = file;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    Active_FilePathTextBox.Text = null;
                    Active_FilePathTextBox.ToolTip = null;
                    break;
            }

            if (result != System.Windows.Forms.DialogResult.Cancel)
                if (!CommonMethods.IsPathRelatedToExcelFile(Active_FilePathTextBox.Text))
                {
                    MessageBox.Show("The file selected is not an excel file!");
                    Active_FilePathTextBox.Text = null;
                    Active_FilePathTextBox.ToolTip = null;
                }
                else
                {
                    
                }
        }
        
        /// <summary>
        /// Selects the the file and puts into the path field for Excluded field.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Excluded_SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new System.Windows.Forms.OpenFileDialog();
            var result = fileDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var file = fileDialog.FileName;
                    Excluded_FilePathTextBox.Text = file;
                    Excluded_FilePathTextBox.ToolTip = file;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    Excluded_FilePathTextBox.Text = null;
                    Excluded_FilePathTextBox.ToolTip = null;
                    break;
            }

            if (result != System.Windows.Forms.DialogResult.Cancel)
                if (!CommonMethods.IsPathRelatedToExcelFile(Excluded_FilePathTextBox.Text))
                {
                    MessageBox.Show("The file selected is not an excel file!");
                    Excluded_FilePathTextBox.Text = null;
                    Excluded_FilePathTextBox.ToolTip = null;
                }
                else
                {
                    CommonMethods.SaveSettingsToConfig();
                }


        }

        /// <summary>
        /// Open a window with a list of users read from the xslt file: Active contenders.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ViewActiveCompsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowWindowWithUsers(Active_FilePathTextBox.Text);
        }

        /// <summary>
        /// Open a window with a list of users read from the xslt file: Excluded contenders.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ViewExcludedCompsButton_Click(object sender, RoutedEventArgs e)
        {
            ShowWindowWithUsers(Excluded_FilePathTextBox.Text);
        }
        
        



        #region Excel loading methods

        /// <summary>
        /// Method Pops-up a window with a list of people selected.
        /// </summary>
        /// <param name="filePath"></param>
        private void ShowWindowWithUsers(string filePath)
        {
            var listOfContenders = CommonMethods.GetContenderList(filePath);

            var window = new Popup();
            window.IsOpen = true;

            

        }


        
                   
               



        #endregion


    }
}
