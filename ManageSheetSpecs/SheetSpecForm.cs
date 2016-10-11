using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Events;

namespace ManageSheetSpecs
{
    public partial class SheetSpecForm : System.Windows.Forms.Form
    {
        public string returnPath { get; set; }
        private BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        private BackgroundWorker backgroundWorker2 = new BackgroundWorker();
        public ConvertWordToPNG convert1 = new ConvertWordToPNG();
        public Image[] images;
        public System.Drawing.Point dimensions;
        public string setWordDocPath;
        public bool run;
        private Document storedDoc;
        public int pages;
        public string sheetSize;
        public double formatHeight;
        public double centerLine;
        public int imagePerSheet;
        public double initialEdgeOffset;
        public double finalYLocation;


        public SheetSpecForm(Document doc, string wordDocPath)
        {
            InitializeComponent();
            SetDoc(doc);
            txtWordPath.Text = wordDocPath;
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            backgroundWorker2.DoWork += new DoWorkEventHandler(backgroundWorker2_DoWork);
            backgroundWorker2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker2_RunWorkerCompleted);

        }

        private void btnSelectWord_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Filter = "Word Documents (*.docx)|*.docx|All files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Multiselect = false;
            ofd.Title = "Select Sheet Specs";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtWordPath.Text = ofd.FileName;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            btnOK.Enabled = false;
            btnSelectWord.Enabled = false;
            SetDocPath(txtWordPath.Text);
            Command.thisCommand.SetFilePath(txtWordPath.Text);
            ConvertingWord();
            this.returnPath = txtWordPath.Text;
            run = true;
            //MakeRequest(RequestId.replaceImages);
            backgroundWorker1.RunWorkerAsync(setWordDocPath);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.returnPath = txtWordPath.Text;
            Close();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string filePath = (string)e.Argument;
            e.Result = convert1.convertWordtoEMF(filePath);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            savingImages();
            backgroundWorker2.RunWorkerAsync(setWordDocPath);
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            string filePath = (string)e.Argument;
            e.Result = convert1.saveEMFasPNG(filePath);
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            savingImages();
            runReplaceImages(setWordDocPath);
            //Document doc = GetDoc();
            //convertWord(storedDoc, setWordDocPath);
            //taskComplete();
        }

        private void runReplaceImages(string wordDocPath)
        {
            string filePath = wordDocPath;
            Document doc = GetDoc();
            int pageCount = getPageCount();
            //convertWord(doc, filePath);
            deleteImages(doc, filePath);
            insertImages(doc, filePath);
            taskComplete();
        }

        public void deleteImages(Document doc, string path)
        {
            string imagePath = Path.GetDirectoryName(setWordDocPath);
            int pageCount = getPageCount();
            SetupProgress(pageCount, "Task: Removing Sheet Specs Images");

            FilteredElementCollector col = new FilteredElementCollector(doc).WhereElementIsNotElementType();
            List<ElementId> ids = new List<ElementId>();

            foreach (Element e in col)
            {
                string imageName = e.Name;
                string typeName = e.GetType().ToString();
                if (typeName == "Autodesk.Revit.DB.Element")
                {
                    if (imageName.StartsWith("Sheet Specs", StringComparison.CurrentCulture))
                    {
                        ids.Add(e.Id);
                    }
                }
            }

            ICollection<ElementId> idsDeleted = null;
            Transaction tx;

            int n = ids.Count;
            if (0 < n)
            {
                using (tx = new Transaction(doc))
                {
                    tx.Start("Delete non-ElementType Sheet Specs Images");
                    idsDeleted = doc.Delete(ids);
                    tx.Commit();
                }
            }

            ids.Clear();

            col = new FilteredElementCollector(doc).WhereElementIsElementType();

            foreach (Element e in col)
            {
                string imageName = e.Name;
                if (imageName.StartsWith("Sheet Specs", StringComparison.CurrentCulture))
                {
                    ids.Add(e.Id);
                }
            }

            n = ids.Count;
            if (0 < n)
            {
                using (tx = new Transaction(doc))
                {
                    tx.Start("Delete ElementType Sheet Specs Images");
                    idsDeleted = doc.Delete(ids);
                    tx.Commit();
                }
            }
        }

        public void insertImages(Document doc, string path)
        {
            string imagePath = Path.GetDirectoryName(path);
            imagePath = imagePath + @"\Sheet Specs (Images)\";
            DirectoryInfo X = new DirectoryInfo(imagePath);
            FileInfo[] someFiles = X.GetFiles("*.png");
            var orderedFiles = someFiles.OrderBy(f => f.FullName);
            FileInfo[] listOfFiles = orderedFiles.ToArray();

            string sheetSize = Command.thisCommand.getSheetSize();

            if (sheetSize == "24 x 36")
            {
                formatHeight = 22.25;
                centerLine = (12.0 / 12);
                imagePerSheet = 4;
                initialEdgeOffset = (3.875 / 12);
                finalYLocation = (23.125 / 12);
            }
            else if (sheetSize == "30 x 42")
            {
                formatHeight = 28.25;
                centerLine = (15.0 / 12);
                imagePerSheet = 5;
                initialEdgeOffset = (3.0 / 12);
                finalYLocation = (29.125 / 12);
            }
            else if (sheetSize == "36 x 48")
            {
                formatHeight = 34.25;
                centerLine = (18.0 / 12);
                imagePerSheet = 6;
                initialEdgeOffset = (2.125 / 12);
                finalYLocation = (35.125 / 12);
            }
            else
            {

            }
            int totalImages = listOfFiles.Count();

            int fullSheets = totalImages / imagePerSheet;
            int lastSheetImageQuantity = totalImages - (imagePerSheet * fullSheets);

            int sheetEndNumber = 01;
            string sheetBegining;
            string templateCategory = Command.thisCommand.getTemplateCategory();
            if (templateCategory == "ARCHITECTURE")
            {
                sheetBegining = "A0.80.";
            }
            else
            {
                sheetBegining = "IA0.80.";
            }
            int imgCount = 0;
            ImageImportOptions iIOptions = new ImageImportOptions();
            iIOptions.Resolution = 150;
            iIOptions.RefPoint = (new XYZ(0, 0, 0));
            iIOptions.Placement = BoxPlacement.TopLeft;
            Autodesk.Revit.DB.View currentImageSheet;
            SetupProgress(listOfFiles.Count(), "Task: Placing Sheet Spec Images");

            //Start with full sheets
            if (fullSheets > 0)
            {
                bool fullSheetsExist = true;
                while (fullSheetsExist)
                {
                    //search for sheet
                    string searchSheet = sheetBegining + sheetEndNumber.ToString("D2");
                    currentImageSheet = FindSheet(doc, searchSheet);
                    if (currentImageSheet == null)
                    {
                        ElementId tBlockID = Command.thisCommand.getTitleBlockID();
                        using (Transaction tx = new Transaction(doc))
                        {
                            tx.Start("Create Sheet");
                            ViewSheet myViewSheet = ViewSheet.Create(doc, tBlockID);
                            myViewSheet.Name = "SPECIFICATIONS";
                            myViewSheet.SheetNumber = searchSheet;
                            tx.Commit();
                        }
                        currentImageSheet = FindSheet(doc, searchSheet);
                        SetSheetParameters(currentImageSheet, doc);
                        double startLocation = initialEdgeOffset;
                        for (int imgOnSheet = 1; imgOnSheet <= imagePerSheet; imgOnSheet++)
                        {
                            if (imgCount < listOfFiles.Count())
                            {
                                iIOptions.RefPoint = (new XYZ(startLocation, finalYLocation, 0));
                                var imageLocation = listOfFiles[imgCount].Directory.FullName + @"\" + listOfFiles[imgCount].Name;
                                Element e = null;
                                using (Transaction tx = new Transaction(doc))
                                {
                                    tx.Start("Import Image");
                                    doc.Import(imageLocation, iIOptions, currentImageSheet, out e);
                                    tx.Commit();
                                }
                                IncrementProgress();
                                startLocation = startLocation + (6.875 / 12);

                                imgCount++;
                            }
                        }
                    }
                    else
                    {
                        double startLocation = initialEdgeOffset;
                        for (int imgOnSheet = 1; imgOnSheet <= imagePerSheet; imgOnSheet++)
                        {
                            if (imgCount < listOfFiles.Count())
                            {
                                iIOptions.RefPoint = (new XYZ(startLocation, finalYLocation, 0));
                                var imageLocation = listOfFiles[imgCount].Directory.FullName + @"\" + listOfFiles[imgCount].Name;
                                Element e = null;
                                using (Transaction tx = new Transaction(doc))
                                {
                                    tx.Start("Import Image");
                                    doc.Import(imageLocation, iIOptions, currentImageSheet, out e);
                                    tx.Commit();
                                }
                                IncrementProgress();
                                startLocation = startLocation + (6.875 / 12);

                                imgCount++;
                            }
                        }
                    }
                    sheetEndNumber++;
                    fullSheets--;
                    if (fullSheets < 1)
                    {
                        fullSheetsExist = false;
                    }
                }
            }

            if (lastSheetImageQuantity > 0)
            {
                int imageInitialOffset = imagePerSheet - lastSheetImageQuantity;
                string searchSheet = sheetBegining + sheetEndNumber.ToString("D2");
                currentImageSheet = FindSheet(doc, searchSheet);
                if (currentImageSheet == null)
                {
                    ElementId tBlockID = Command.thisCommand.getTitleBlockID();
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Create Sheet");
                        ViewSheet myViewSheet = ViewSheet.Create(doc, tBlockID);
                        myViewSheet.Name = "SPECIFICATIONS";
                        myViewSheet.SheetNumber = searchSheet;
                        tx.Commit();
                    }
                    currentImageSheet = FindSheet(doc, searchSheet);
                    SetSheetParameters(currentImageSheet, doc);
                    double startLocation = initialEdgeOffset + (imageInitialOffset * (6.875 / 12));
                    for (int imgOnSheet = imageInitialOffset; imgOnSheet <= imagePerSheet; imgOnSheet++)
                    {
                        if (imgCount < listOfFiles.Count())
                        {
                            iIOptions.RefPoint = (new XYZ(startLocation, finalYLocation, 0));
                            var imageLocation = listOfFiles[imgCount].Directory.FullName + @"\" + listOfFiles[imgCount].Name;
                            Element e = null;
                            using (Transaction tx = new Transaction(doc))
                            {
                                tx.Start("Import Image");
                                doc.Import(imageLocation, iIOptions, currentImageSheet, out e);
                                tx.Commit();
                            }
                            IncrementProgress();
                            startLocation = startLocation + (6.875 / 12);

                            imgCount++;
                        }
                    }
                }
                else
                {
                    double startLocation = initialEdgeOffset + (imageInitialOffset * (6.875 / 12));
                    for (int imgOnSheet = imageInitialOffset; imgOnSheet <= imagePerSheet; imgOnSheet++)
                    {
                        if (imgCount < listOfFiles.Count())
                        {
                            iIOptions.RefPoint = (new XYZ(startLocation, finalYLocation, 0));
                            var imageLocation = listOfFiles[imgCount].Directory.FullName + @"\" + listOfFiles[imgCount].Name;
                            Element e = null;
                            using (Transaction tx = new Transaction(doc))
                            {
                                tx.Start("Import Image");
                                doc.Import(imageLocation, iIOptions, currentImageSheet, out e);
                                tx.Commit();
                            }
                            IncrementProgress();
                            startLocation = startLocation + (6.875 / 12);

                            imgCount++;
                        }
                    }
                }
            }
        }

        public void SetSheetParameters(Autodesk.Revit.DB.View view, Document doc)
        {
            string templateCategory = Command.thisCommand.getTemplateCategory();
            ViewSheet mySheet = view as ViewSheet;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Organize Sheet");
                mySheet.LookupParameter("*Discipline").Set("GENERAL");
                mySheet.LookupParameter("*Discipline Code").Set("00");
                if (templateCategory == "ARCHITECTURE")
                {
                    mySheet.LookupParameter("*Discipline Subcode").Set("A080 SPECIFICATIONS");
                }
                else
                {
                    mySheet.LookupParameter("*Discipline Subcode").Set("IA080 SPECIFICATIONS");
                }
                tx.Commit();
            }

        }

        public Autodesk.Revit.DB.View FindSheet(Document doc, string searchSheet)
        {
            Autodesk.Revit.DB.View rtnView = null;
            FilteredElementCollector col = new FilteredElementCollector(doc);
            col.OfCategory(BuiltInCategory.OST_Sheets);
            col.OfClass(typeof(ViewSheet));

            foreach (Element v in col)
            {
                ViewSheet vs = v as ViewSheet;
                var sheetNumber = vs.SheetNumber;
                if (sheetNumber == searchSheet)
                {
                    rtnView = vs as Autodesk.Revit.DB.View;
                    return rtnView;
                }
            }
            return rtnView;
        }

        /*public void convertWord(Document doc, string path)
        {
            string imagePath = Path.GetDirectoryName(setWordDocPath);
            int pageCount = getPageCount(); 
            SetupProgress(pageCount, "Task: Replacing Sheet Spec Images");


            FilteredElementCollector col = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RasterImages);

            int counter = 1;
            foreach (Element e in col)
            {
                var tempTest = e.GetType();
                if (tempTest.FullName == "Autodesk.Revit.DB.ImageType")
                {
                    string imageName = e.Name;
                    int index = imageName.LastIndexOf(" ");
                    if (index > 5)
                        imageName = imageName.Substring(0, index);
                    string fullImagePath = imagePath + @"\Sheet Specs (Images)\" + imageName;
                    if (imageName.StartsWith("Sheet Specs", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (File.Exists(fullImagePath))
                        {
                            using (Transaction tx = new Transaction(doc))
                            {
                                tx.Start("Replace Image");
                                ImageType image = e as ImageType;
                                image.ReloadFrom(fullImagePath);
                                tx.Commit();
                            }
                            ++counter;
                            if (counter < pageCount)
                            {
                                IncrementProgress();
                            }
                        }
                    }
                }
            }

        }*/

        #region set and get doc path
        private void SetDocPath(string wordDocPath)
        {
            setWordDocPath = wordDocPath;
        }

        public string GetDocPath()
        {
            return setWordDocPath;
        }

        private void SetDoc(Document doc)
        {
            storedDoc = doc;
        }

        public Document GetDoc()
        {
            return storedDoc;
        }

        public void setPageCount(int pageCount)
        {
            pages = pageCount;
        }

        public int getPageCount()
        {
            return pages;
        }
        #endregion

        private void ConvertingWord()
        {
            lblProcessing1.Text = "Task: Processing Word Document... Please Wait";
            lblProcessing1.Visible = true;
            pictureBox1.Visible = true;
        }

        public void SetupProgress(int max, string info)
        {
            MethodInvoker mi = delegate
            {
                progressBar1.Minimum = 0;
                progressBar1.Maximum = max;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                lblProcessing1.Text = info;
                this.Refresh();
            };
            if (InvokeRequired)
            {
                this.Invoke(mi);
            }
            else
            {
                progressBar1.Minimum = 0;
                progressBar1.Maximum = max;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                lblProcessing1.Text = info;
                this.Refresh();
            }            
        }

        public void IncrementProgress()
        {
            MethodInvoker mi = delegate
            {
                ++progressBar1.Value;
                this.Refresh();
            };
            if (InvokeRequired)
            {
                this.Invoke(mi);
            }
            else
            {
                ++progressBar1.Value;
                this.Refresh();
            }
            System.Windows.Forms.Application.DoEvents();
        }

        public void taskComplete()
        {
            MethodInvoker mi = delegate
            {
                lblProcessing1.Text = "Task Complete: Sheet Specs have been updated.";
                progressBar1.Visible = false;
                this.Refresh();
            };
            if (InvokeRequired)
            {
                this.Invoke(mi);
            }
            else
            {
                lblProcessing1.Text = "Task Complete: Sheet Specs have been updated.";
                progressBar1.Visible = false;
                this.Refresh();
            }
        }

        private void savingImages()
        {
            MethodInvoker mi = delegate
            {
                pictureBox1.Visible = false;
                lblProcessing1.Text = "Task: Converting Word Document to Images";
                lblProcessing1.Visible = true;
                this.Refresh();
            };
            if (InvokeRequired)
            {
                this.Invoke(mi);
            }
            else
            {
                pictureBox1.Visible = false;
                lblProcessing1.Text = "Task: Converting Word Document to Images";
                lblProcessing1.Visible = true;
                this.Refresh();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            this.returnPath = txtWordPath.Text;
            this.DialogResult = DialogResult.OK;
            base.OnFormClosing(e);
        }
    }
}
