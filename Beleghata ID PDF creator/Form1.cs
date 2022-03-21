using System;
using System.Drawing;
using System.Windows.Forms;
using NovaNet.Utils;
using System.Data;
using System.Data.Odbc;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Data.SqlClient;
using System.Collections.ObjectModel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ClearImage;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.OCRProcessor;
using Syncfusion.OCRProcessor.Interop;

namespace Beleghata_ID_PDF_creator
{
    public partial class Form1 : Form
    {
        ListViewItem lvi;
        private Dictionary<string, ListViewItem> ListViewItems = new Dictionary<string, ListViewItem>();
        private Dictionary<string, ListViewItem> ListViewItems1 = new Dictionary<string, ListViewItem>();

        string[] imageName;
        string imageName1;
        string filespath;

        public static string foldername;
        public Imagery img;
        public string filename;
        public NovaNet.Utils.Imagery img1;

        public Form1()
        {
            InitializeComponent();
            img = IgrFactory.GetImagery(Constants.IGR_CLEARIMAGE);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            img = IgrFactory.GetImagery(Constants.IGR_CLEARIMAGE);

            //toolStripStatusLabel1.Visible = true;
            //toolStripStatusLabel1.Text = "Select Specific Image Folder...";
            //toolStripProgressBar1.Visible = false;
            txtPath.Enabled = false;
            btnPDF.Enabled = false;
            
            btnBrowse.Select();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            List<string> fileNames = new List<string>();
            List<string> tempPath = new System.Collections.Generic.List<string>(1000);
            DialogResult dr = folderBrowserDialog1.ShowDialog();

            //toolStripStatusLabel1.Text = "Select Specific Image Folder...";

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                txtPath.Text = folderBrowserDialog1.SelectedPath;

                DirectoryInfo selectedPath = new DirectoryInfo(txtPath.Text);
                ListViewItems.Clear();
                ListViewItems1.Clear();
                lstImage.Items.Clear();

                if (selectedPath.GetDirectories().Length > 0)
                {
                    string[] folders = Directory.GetDirectories(txtPath.Text);

                    for (int i = 0; i < folders.Length; i++)
                    {
                        //MessageBox.Show(folders[i].ToString());
                        lstImage.Items.Add(System.IO.Path.GetFileName(folders[i].ToString()));
                        ListViewItems.Add(folders[i].ToString(), lvi);
                    }
                }

                   /* foreach (FileInfo file in selectedPath.GetFiles())
                    {
                        if (file.Extension.Equals(".jpg") || file.Extension.Equals(".JPG") || (file.Extension.Equals(".JPEG")) || (file.Extension.Equals(".jpeg")) || file.Extension.Equals(".tif") || file.Extension.Equals(".TIF") || file.Extension.Equals(".pdf"))
                        {
                            fileNames.Add(file.FullName);
                            tempPath.Add(txtPath.Text + "\\" + file.ToString());

                        }
                    }*/
                /*ListViewItems.Clear();
                ListViewItems1.Clear();
                lstImage.Items.Clear();*/

                foldername = selectedPath.Name;
                
                if (lstImage.Items.Count > 0)
                {
                    btnPDF.Enabled = true;
                    //btnMerge.Enabled = true;
                    //toolStripStatusLabel1.Text = "Folder is selected successfully...PDF is ready to create \t";
                    //toolStripProgressBar1.Visible = false;
                }

            }
            else
            {
                btnPDF.Enabled = false;
                //toolStripStatusLabel1.Visible = false;
                //toolStripProgressBar1.Visible = false;
                //toolStripStatusLabel1.Text = "";
                return;
            }
        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            List<string> fileNames = new List<string>();
            List<string> tempPath = new System.Collections.Generic.List<string>(1000);
            try
            {
                

                //toolStripStatusLabel1.Text = "PDF is creating... \t";

                btnBrowse.Enabled = false;
                btnPDF.Enabled = false;
               

                imageName = null;


                string expFolder = "C:\\";
                bool isDeleted = false;

                //toolStripProgressBar1.Value = 5;
                if (lstImage.Items.Count > 0)
                {
                    ListViewItems.Clear();
                    ListViewItems1.Clear();
                    listBox1.Items.Clear();
                    

                    //toolStripStatusLabel1.Text = "PDF is creating... \t";
                    
                    for (int a = 0; a < lstImage.Items.Count; a++)
                    {
                        
                        string filename = lstImage.Items[a].ToString();
                        //imageName[a] = dsexport.Tables[0].Rows[x][4].ToString() + "\\QC" + "\\" + dsimage.Tables[0].Rows[a]["page_name"].ToString();

                        //imageName[a] = txtPath.Text + "\\" + filename.ToString();
                        filespath = txtPath.Text + "\\" + filename.ToString();
                        DirectoryInfo selectedPath = new DirectoryInfo(filespath);
                        
                        foreach (FileInfo file in selectedPath.GetFiles())
                        {
                            if (file.Extension.Equals(".jpg") || file.Extension.Equals(".JPG") || (file.Extension.Equals(".JPEG")) || (file.Extension.Equals(".jpeg")) || file.Extension.Equals(".tif") || file.Extension.Equals(".TIF") || file.Extension.Equals(".pdf"))
                            {
                                //fileNames.Add(file.FullName);
                                tempPath.Add(filespath + "\\" + file.ToString());

                                //listBox1.Items.Add(System.IO.Path.GetFileName(file.FullName));
                                //ListViewItems1.Add(file.FullName, lvi);
                                listBox1.Items.Add(System.IO.Path.GetFileName(file.FullName));
                                //lstImage.Items.Add(fileName);
                                //ListViewItem lvi1 = lstTotalImage.Items.Add(System.IO.Path.GetFileNameWithoutExtension(fileName));
                                //lvi.Tag = fileName;
                                //lvi1.Tag = fileName;
                                ListViewItems.Add(file.FullName, lvi);
                            }
                            
                        }
                        
                        imageName = new string[listBox1.Items.Count];
                        for (int b = 0; b < listBox1.Items.Count; b++)
                        {
                            imageName[b] = filespath + "\\" + listBox1.Items[b].ToString();
                        }
                        
                        if (imageName.Length != 0)
                        {
                            expFolder = "C:\\";
                            //toolStripStatusLabel1.Text = "PDF is creating... \t";
                            if (Directory.Exists(expFolder + "\\Export\\" + foldername + "\\" + filename) && isDeleted == false)
                            {
                                Directory.Delete(expFolder + "\\Export\\" + foldername + "\\" + filename, true);

                            }

                            if (!Directory.Exists(expFolder + "\\Export\\" + foldername + "\\" + filename))
                            {
                                Directory.CreateDirectory(expFolder + "\\Export\\" + foldername + "\\" + filename);
                                isDeleted = true;
                            }
                            expFolder = "C:\\";

                            if (listBox1.Items.Count > 0)
                            {
                                //toolStripStatusLabel1.Text = "PDF is creating... \t";


                                if (img.TifToPdf(imageName, 80, expFolder + "\\Export\\" + foldername + "\\" + filename + "\\" + filename + ".pdf") == true)
                                {

                                    //toolStripProgressBar1.Visible = true;
                                    //toolStripProgressBar1.Value = 100;
                                    //toolStripStatusLabel1.Visible = true;
                                    //toolStripStatusLabel1.Text = "PDF created Successfully... \t";
                                    string pdf_path = expFolder + "\\Export\\" + foldername + "\\" + filename + "\\" + filename + ".pdf";
                                    string file = filename + ".pdf";
                                    string dirname = Path.GetDirectoryName(pdf_path);
                                    //PdfDocument document = new PdfDocument(PdfPage.PAGE.ToPdf,PdfImage.BEST_COMPRESSION);
                                    PdfReader pdfReader = new PdfReader(pdf_path);
                                    int noofpages = pdfReader.NumberOfPages;

                                    List<string> fileNamesNew = new List<string>();

                                    iTextSharp.text.Document document = new iTextSharp.text.Document();

                                    //ocr directory create
                                    string dirEx = dirname + "\\OCR";
                                    if (!Directory.Exists(dirEx))
                                    {
                                        Directory.CreateDirectory(dirEx);
                                    }

                                    //split pdf
                                    for (int i = 0; i < noofpages; i++)
                                    {
                                        using (MemoryStream ms = new MemoryStream())
                                        {
                                            PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdf_path);
                                            Syncfusion.Pdf.PdfDocument documentPage = new Syncfusion.Pdf.PdfDocument();
                                            documentPage.ImportPage(loadedDocument, i);
                                            documentPage.Save(dirEx + "\\OCR_" + i + ".pdf");
                                            string filenameNew = dirEx + "\\OCR_" + i + ".pdf";
                                            //documentPage.Close();
                                            documentPage.Close(true);
                                            //documentPage.Dispose();
                                            //loadedDocument.Close();
                                            loadedDocument.Close(true);
                                            //loadedDocument.Dispose();
                                            documentPage.EnableMemoryOptimization = true;
                                            loadedDocument.EnableMemoryOptimization = true;
                                            fileNamesNew.Add(filenameNew);
                                            //lstImage.Items.Add(filenameNew);

                                            ms.Close();

                                            GC.Collect();
                                            GC.WaitForPendingFinalizers();
                                            GC.Collect();

                                            Application.DoEvents();
                                        }
                                        Application.DoEvents();
                                    }

                                    string expath = Path.GetDirectoryName(Application.ExecutablePath);
                                    //ocr
                                    try
                                    {
                                        System.IO.DirectoryInfo di3 = new DirectoryInfo(dirEx);
                                        foreach (string filename1 in fileNamesNew)
                                        {
                                            Application.DoEvents();
                                            string xyz = filename1;
                                            //PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdf_path);
                                            //Syncfusion.Pdf.PdfDocument documentPage = new Syncfusion.Pdf.PdfDocument();
                                            //documentPage.ImportPage(loadedDocument, i);
                                            using (OCRProcessor oCR = new OCRProcessor(expath + "\\TesseractBinaries\\3.02\\"))
                                            {
                                                try
                                                {

                                                    PdfLoadedDocument pdfLoadedDocument = new PdfLoadedDocument(xyz);

                                                    oCR.Settings.Language = Syncfusion.OCRProcessor.Languages.English;

                                                    oCR.PerformOCR(pdfLoadedDocument, expath + "\\tessdata\\", true);

                                                    pdfLoadedDocument.EnableMemoryOptimization = true;

                                                    pdfLoadedDocument.Save(xyz);

                                                    pdfLoadedDocument.Close(true);

                                                    oCR.Dispose();

                                                    GC.Collect();
                                                    GC.WaitForPendingFinalizers();
                                                    GC.Collect();
                                                }
                                                catch(Exception)
                                                { continue; }
                                            }

                                        }
                                        string outFile = expFolder + "\\Export\\" + foldername + "\\" + filename + "\\" + filename + ".pdf";
                                        try
                                        {
                                            //create newFileStream object which will be disposed at the end
                                            using (FileStream newFileStream = new FileStream(outFile, FileMode.Create))
                                            {
                                                Application.DoEvents();
                                                // step 2: we create a writer that listens to the document
                                                PdfCopy writer = new PdfCopy(document, newFileStream);
                                                if (writer == null)
                                                {
                                                    return;
                                                }

                                                // step 3: open the document
                                                document.Open();

                                                foreach (string filename1 in fileNamesNew)
                                                {
                                                    Application.DoEvents();
                                                    string xyz = filename1;
                                                    // create a reader for a certain document
                                                    PdfReader reader = new PdfReader(xyz);
                                                    reader.ConsolidateNamedDestinations();

                                                    // step 4: add content
                                                    for (int i = 1; i <= reader.NumberOfPages; i++)
                                                    {
                                                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                                                        writer.AddPage(page);
                                                    }

                                                    PRAcroForm form = reader.AcroForm;
                                                    if (form != null)
                                                    {
                                                        writer.CopyAcroForm(reader);
                                                    }

                                                    reader.Close();
                                                }

                                                // step 5: close the document and writer
                                                writer.Close();
                                                document.Close();

                                                GC.Collect();
                                                GC.WaitForPendingFinalizers();
                                                GC.Collect();
                                            }//disposes the newFileStream object

                                            Directory.Delete(dirEx, true);

                                            //MessageBox.Show("OCR Completed Successfully ...");
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.ToString());
                                        }
                                    }
                                    catch (Exception ex1)
                                    { MessageBox.Show(ex1.ToString()); }

                                }
                                else
                                {
                            
                                    //toolStripProgressBar1.Value = 0;
                                    //toolStripStatusLabel1.Visible = true;
                                    //toolStripStatusLabel1.Text = "Error Occured \t";
                                    MessageBox.Show("There is a problem in one or more pages of PDF name: " + foldername + ".pdf"+ "\n The error is: " + img.GetError());
                                    return;
                                }
                                
                            }

                            ListViewItems.Clear();
                            ListViewItems1.Clear();
                            listBox1.Items.Clear();
                            

                        }

                        
                    }
                    btnBrowse.Enabled = true;
                    MessageBox.Show("Pdf is created successfully !");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
