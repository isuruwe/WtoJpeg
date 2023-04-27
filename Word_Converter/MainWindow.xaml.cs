using System;
using System.IO;
using Microsoft.Win32;
using System.Windows;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
//using Microsoft.Office.Interop.Word;

namespace Word_Converter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.Multiselect = true;
            //openFileDialog.Filter = "Text files (*.pdf)|*.pdf|All files (*.*)|*.*";
            //openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //if (openFileDialog.ShowDialog() == true)
            //{
            //    foreach (string filename in openFileDialog.FileNames)
            //        lbFiles.Items.Add(filename);
            //        //lbFiles.Items.Add(Path.GetFileName(filename));
            //}

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {


            Microsoft.Office.Interop.Word.Application myWordApp = new Microsoft.Office.Interop.Word.Application();
            Document myWordDoc = new Document();
            object missing = System.Type.Missing;
            string pdfPath = @"e:\EJAET-3-2-45-47.pdf";
            foreach (var item in lblWord.Items)
            {
                pdfPath = item.ToString();
            }
            object path1 = pdfPath;
            myWordDoc = myWordApp.Documents.Add(path1, missing, missing, missing);
            myWordApp.Visible = true;
            foreach (Microsoft.Office.Interop.Word.Window window in myWordDoc.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var bits = pane.Pages[i].EnhMetaFileBits;
                        var target = path1 + i.ToString() + "_image.doc";
                        try
                        {
                            using (var ms = new MemoryStream((byte[])(bits)))
                            {


                                var image = System.Drawing.Image.FromStream(ms);
                                var pngTarget = Path.ChangeExtension(target, "jpeg");
                                using (var b = new Bitmap(image.Width, image.Height))
                                {
                                    b.SetResolution(image.HorizontalResolution, image.VerticalResolution);

                                    using (var g = Graphics.FromImage(b))
                                    {
                                        g.Clear(Color.White);
                                        g.DrawImageUnscaled(image, 0, 0);
                                    }

                                    b.Save(pngTarget, System.Drawing.Imaging.ImageFormat.Jpeg);
                                }

                            }
                        }
                        catch (System.Exception ex)
                        { }
                    }
                }
            }
            object ojnf = false; 
            myWordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            myWordApp.Quit(ref ojnf, ref ojnf, ref ojnf);
           
            myWordApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(myWordApp);
            myWordApp = null;
            string prossnme = "winword";
            Process[] prssess = Process.GetProcessesByName(prossnme);
            foreach(Process proses in prssess)
            {
                proses.Kill();
            }

            
            
            MessageBox.Show("Convert Success!");

            ////Convert PDF into specified Image height & width
            //SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            ////this property is necessary only for registered version
            //f.Serial = "1234567890";

            //// Set initial values
            //string pdfPath = @"d:\EJAET-3-2-45-47.pdf";
            //foreach (var item in lbFiles.Items)
            //{
            //    pdfPath = item.ToString();
            //}

            //string imageFolder = Path.GetDirectoryName(pdfPath);
            //int width = 1600; // Width in Px
            //int height = 1900; // Height in Px

            ////Set image options
            //f.ImageOptions.ImageFormat = ImageFormat.Png;
            //f.ImageOptions.Resize(new System.Drawing.Size { Width = width, Height = height }, false);


            //f.OpenPdf(pdfPath);
            //if (f.PageCount > 0)
            //{
            //    // Convert all pages to PNG images
            //    f.ToImage(imageFolder, "Page");

            //    //Show image
            //    System.Diagnostics.Process.Start(imageFolder);

            //}

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Text files (*.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string filename in openFileDialog.FileNames)
                    lblWord.Items.Add(filename);
                //lbFiles.Items.Add(Path.GetFileName(filename));
            }
           


            //string pdfPath = @"e:\EJAET-3-2-45-47.pdf";
            //foreach (var item in lblWord.Items)
            //{
            //    pdfPath = item.ToString();
            //}

            //var app = new Microsoft.Office.Interop.Word.Application();
            //var doc = app.Documents.Open(pdfPath);

            ////Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            ////Microsoft.Office.Interop.Word.Document oDocument=new Microsoft.Office.Interop.Word.Document();
            ////oWord

            ////Opens the word document and fetch each page and converts to image
            //foreach (Microsoft.Office.Interop.Word.Window window in doc.Windows)
            //{
            //    foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
            //    {
            //        for (var i = 1; i <= pane.Pages.Count; i++)
            //        {
            //            var page = pane.Pages[i];

            //            var bits = page.EnhMetaFileBits;
            //            var target = pdfPath;

            //            try
            //            {
            //                using (var ms = new MemoryStream((byte[])(bits)))
            //                {

            //                    var image = System.Drawing.Image.FromStream(ms);                                
            //                    var pngTarget = Path.ChangeExtension(target, "jpeg");
            //                    image.Save(pngTarget, ImageFormat.Bmp);
            //                }
            //            }
            //            catch (System.Exception ex)
            //            { }
            //        }
            //    }
            //}
            //doc.Close(Type.Missing, Type.Missing, Type.Missing);

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            lblWord.Items.Clear();
        }
    }
}
