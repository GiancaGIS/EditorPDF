using Microsoft.Win32;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace EditorPDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<int> _pagineDaEliminare = null;
        private FileInfo _fileInfo = null;

        public MainWindow()
        {
            _pagineDaEliminare = new List<int>();
            InitializeComponent();
        }

        private void MergeMultiplePDFIntoSinglePDF(string outputFilePath, string[] pdfFiles)
        {
            PdfDocument document = new PdfDocument();
            foreach (string pdfFile in pdfFiles)
            {
                PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                document.Version = inputPDFDocument.Version;
                foreach (PdfPage page in inputPDFDocument.Pages)
                {
                    document.AddPage(page);
                }
                // When document is add in pdf document remove file from folder  
                System.IO.File.Delete(pdfFile);
            }
            // Set font for paging  
            XFont font = new XFont("Verdana", 9);
            XBrush brush = XBrushes.Black;
            // Create variable that store page count  
            string noPages = document.Pages.Count.ToString();
            // Set for loop of document page count and set page number using DrawString function of PdfSharp  
            for (int i = 0; i < document.Pages.Count; ++i)
            {
                PdfPage page = document.Pages[i];
                // Make a layout rectangle.  
                XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.DrawString("Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);
                }
            }
            document.Options.CompressContentStreams = true;
            document.Options.NoCompression = false;
            // In the final stage, all documents are merged and save in your output file path.  
            document.Save(outputFilePath);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_fileInfo is null) throw new Exception("Selezionare un file PDF valido!");

                if (!_fileInfo.Extension.Equals(".pdf", StringComparison.InvariantCultureIgnoreCase)) throw new Exception("Selezionare un file PDF valido!");

                EngineDelete();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Attenzione", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void EngineDelete()
        {
            await Task.Run(() =>
            {
                string nomeFile = _fileInfo.Name;
                string directory = _fileInfo.DirectoryName;

                PdfDocument document = new PdfDocument();

                using (PdfDocument inputPDFDocument = PdfReader.Open(_fileInfo.FullName, PdfDocumentOpenMode.Import))
                {
                    PdfPages pagine = inputPDFDocument.Pages;
                    int contatorePagine = 1;

                    foreach (PdfPage pagina in pagine)
                    {
                        if (!_pagineDaEliminare.Contains(contatorePagine))
                            _ = document.AddPage(pagina);

                        contatorePagine += 1;
                    }

                    document.Options.CompressContentStreams = true;
                    document.Options.NoCompression = false;
                    // In the final stage, all documents are merged and save in your output file path.
                    string pathFileOutput = Path.Combine(directory,
                        $"{nomeFile.Replace($"{_fileInfo.Extension}", "")}_editato_{DateTime.Now.ToShortDateString().Replace("/", "_")}.pdf");

                    document.Save(pathFileOutput);

                    document.Close();
                    document.Dispose();
                    MessageBoxResult dialogResult = MessageBox.Show($"Fine elaborazione.{Environment.NewLine}Aprire il file?", "Tutto ok", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start(pathFileOutput);
                    }
                }
            });
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            _pagineDaEliminare.Clear();
            string testoUtente = TextBoxPagine.Text;
            List<string> primoSottoInsieme = testoUtente.Split(new char[] { ' ', ',', '-' }).ToList();
            foreach (string primeSottoStringhe in primoSottoInsieme)
            {
                if (int.TryParse(primeSottoStringhe, out int paginaDaEiminare))
                {
                    _pagineDaEliminare.Add(paginaDaEiminare);
                }

            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog selez_file = new OpenFileDialog()
            {
                Filter = "Text files (*.pdf)|*.pdf|All files (*.*)|*.*"
            };

            bool? sceltaUtente = selez_file.ShowDialog();

            if (sceltaUtente.HasValue && sceltaUtente.Value)
            {
                _fileInfo = new FileInfo(selez_file.FileName);
                TextBoxPathPDFScelto.Text = selez_file.FileName; // Imposto il solo nome del file selezionato alla prima textbox
                //TextBoxPathPDFScelto.Text = selez_file.SafeFileName; // imposto tutto il percorso!
            }
        }
    }
}
