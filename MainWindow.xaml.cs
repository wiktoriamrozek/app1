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

using Microsoft.Win32;

using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.SfSkinManager;
using Syncfusion.Themes.FluentLight.WPF;

using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            FluentTheme fluentTheme = new FluentTheme()
            {
                ThemeName = "FluentLight",
                HoverEffectMode = HoverEffect.None,
                PressedEffectMode = PressedEffect.Glow,
                ShowAcrylicBackground = false
            };

            FluentLightThemeSettings themeSettinggs = new FluentLightThemeSettings();
            themeSettinggs.BodyFontSize = 16;
            themeSettinggs.FontFamily = new System.Windows.Media.FontFamily("Barlow");
            SfSkinManager.RegisterThemeSettings("FluentLight", themeSettinggs);
            SfSkinManager.SetTheme(this, fluentTheme);

            InitializeComponent();
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if(pathTexBox.Text == String.Empty)
            {
                MessageBox.Show("Please select a file");
                return;
            }

            switch(conversionDropDown.SelectedIndex)
            {
                case 0:
                    ConvertDocToPDF(pathTexBox.Text);
                    break;
                case 1:
                    ConvertPDFToDoc(pathTexBox.Text);
                    break;
                case 2:
                    ConvertPNGToPDF(pathTexBox.Text);
                    break;
                default:
                    MessageBox.Show("Please select an option");
                    return;
            }

            OpenFolder(pathTexBox.Text);
        }

        private void ConvertDocToPDF(string docPath)
        {
            WordDocument wordDocument = new WordDocument(docPath, FormatType.Automatic);
            DocToPDFConverter converter = new DocToPDFConverter();
            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);

            string newPDFPath = docPath.Split('.')[0] + ".pdf";
            pdfDocument.Save(newPDFPath);

            pdfDocument.Close(true);
            wordDocument.Close();
        }

        private void ConvertPNGToPDF(string pngPath)
        {
            PdfDocument pdfDoc = new PdfDocument();
            PdfImage pdfimage = PdfImage.FromStream(new FileStream(pngPath, FileMode.Open));
            PdfPage pdfPage = new PdfPage();
            PdfSection pdfSection = pdfDoc.Sections.Add();

            pdfSection.Pages.Insert(0, pdfPage);
            pdfPage.Graphics.DrawImage(pdfimage, 0, 0);

            string newPNGPath = pngPath.Split('.')[0] + ".pdf";
            pdfDoc.Save(newPNGPath);
            pdfDoc.Close(true);
        }

        private void ConvertPDFToDoc(string pdfPath)
        {
            WordDocument wordDocument = new WordDocument();
            IWSection section = wordDocument.AddSection();
            section.PageSetup.Margins.All = 0;
            IWParagraph firstParagraph = section.AddParagraph();

            SizeF defaltPageSize = new SizeF(wordDocument.LastSection.PageSetup.PageSize.Width,
                wordDocument.LastSection.PageSetup.PageSize.Height);

            using(PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdfPath))
            {
                for (int i = 0; i < loadedDocument.Pages.Count; i++)
                {
                    using(var image = loadedDocument.ExportAsImage(i, defaltPageSize, false))
                    {
                        IWPicture picture = firstParagraph.AppendPicture(image);
                        picture.Width = image.Width;
                        picture.Height = image.Height;
                    }
                }
            };

            string newPDFPath = pdfPath.Split('.')[0] + ".docx";
            wordDocument.Save(newPDFPath);
            wordDocument.Dispose();
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            bool? result = openFileDialog.ShowDialog();

            if(result.HasValue && result.Value)
            {
                pathTexBox.Text = openFileDialog.FileName;
            }
        }

        private void OpenFolder(string folderPath) 
        {
            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                Arguments = folderPath.Substring(0, folderPath.LastIndexOf('\\')),
                FileName = "explorer.exe"
            };

            Process.Start(startInfo);
        }

        private void myMainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            {
                myMainWindow.Width = e.NewSize.Width;
                myMainWindow.Height = e.NewSize.Height;

                double xChange = 1, yChange = 1;

                if (e.PreviousSize.Width != 0)
                    xChange = (e.NewSize.Width / e.PreviousSize.Width);

                if (e.PreviousSize.Height != 0)
                    yChange = (e.NewSize.Height / e.PreviousSize.Height);

                foreach (FrameworkElement fe in myGrid.Children)
                {
                    if (fe is Grid == false)
                    {
                        fe.Height = fe.ActualHeight * yChange;
                        fe.Width = fe.ActualWidth * xChange;

                        Canvas.SetTop(fe, Canvas.GetTop(fe) * yChange);
                        Canvas.SetLeft(fe, Canvas.GetLeft(fe) * xChange);

                    }
                }
            }
        }
    }

}
