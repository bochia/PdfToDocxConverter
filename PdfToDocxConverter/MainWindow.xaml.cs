using Microsoft.Win32;
using System.Windows;

namespace PdfToDocxConverter
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

        private void btnSelectInputFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "PDF files |*.pdf"};
            openFileDialog.ShowDialog();

            var directoryPath = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
            var fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(openFileDialog.SafeFileName);
            var wordDocxFilePath = directoryPath + "\\" + fileNameWithoutExtension + ".docx";


            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            f.OpenPdf(openFileDialog.FileName);

            if (f.PageCount > 0)
            {
                f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;
                f.ToWord(wordDocxFilePath);
            }
        }

    }
}
