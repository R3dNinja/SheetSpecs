using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace ManageSheetSpecs
{
    public class ConvertWordToPNG
    {
        public delegate void ProgressUpdate(int value);
        public event ProgressUpdate OnProgressUpdate;

        public string wordDocPath;
        private int bitmapWidth;
        private int bitmapHeight;
        public string sheetSize;
        public float formatHeight;
        public double centerLine;
        public int imagePerSheet;

        public Image[] convertWordtoEMF(string path)
        {
            double sWidth = Command.thisCommand.getSheetWidth();
            setSheetParameters(sWidth);
            string sheetSize = Command.thisCommand.getSheetSize();

            if (sheetSize == "24 x 36")
            {
                formatHeight = 1602; //22.25;
                centerLine = 12.0;
                imagePerSheet = 4;
            }
            else if (sheetSize == "30 x 42")
            {
                formatHeight = 2034; //28.25;
                centerLine = 15.0;
                imagePerSheet = 5;
            }
            else if (sheetSize == "36 x 48")
            {
                formatHeight = 2466; //34.25;
                centerLine = 18.0;
                imagePerSheet = 6;
            }
            else
            {

            }
            //set resolution for export and empty page size
            int resolution = 150;
            float pWidth = 0.00f;
            float pHeight = 0.00f;

            MarginsF pagemargins = new MarginsF();
            pagemargins.Bottom = 0;
            pagemargins.Top = 0;
            pagemargins.Left = 0;
            pagemargins.Right = 1;

            //Load the document
            WordDocument doc = new WordDocument();
            doc.OpenReadOnly(path, FormatType.Docx);

            setWordPath(path);

            //Get the Document size
            foreach (WSection section in doc.Sections)
            {
                section.PageSetup.Margins = pagemargins;
                section.PageSetup.PageSize = new SizeF(477, formatHeight);
                pHeight = (section.PageSetup.PageSize.Height) / 72;
                pWidth = (section.PageSetup.PageSize.Width) / 72;
            }

            //calculate image size for bitmap
            bitmapWidth = Convert.ToInt32(pWidth * resolution);
            bitmapHeight = Convert.ToInt32(pHeight * resolution);

            Point dimension = new Point();
            dimension.X = bitmapWidth;
            dimension.Y = bitmapHeight;

            //convert doc to Image
            Image[] images = doc.RenderAsImages(ImageType.Metafile);
            doc.Close();

            Command.thisCommand.holdImages(images);
            Command.thisCommand.setDimension(dimension);
            return images;
        }

        public int saveEMFasPNG(string path)
        {
            Image[] images = Command.thisCommand.getImages();
            Point dimension = Command.thisCommand.getDimension();

            bitmapWidth = dimension.X;
            bitmapHeight = dimension.Y;
            //string path = getWordPath();
            var savePath = Path.GetDirectoryName(path) + @"\Sheet Specs (Images)\";
            if (Directory.Exists(savePath))
            {
                removeExistingImages(savePath);
            }
            else
            {
                Directory.CreateDirectory(savePath);
            }
            var saveName = "Sheet Specs";
            int resolution = 150;
            int pageCount = images.Count();
            int i = 1;

            Command.thisCommand.dialog.setPageCount(pageCount);
            Command.thisCommand.dialog.SetupProgress(pageCount, "Task: Converting Word Document to Images");



            foreach (Image image in images)
            {
                Bitmap bitmap = null;
                Metafile metafile = image as Metafile;
                bitmap = new Bitmap(bitmapWidth, bitmapHeight);
                bitmap.SetResolution((float)resolution, (float)resolution);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
                    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    g.DrawImage(metafile, 0, 0, (float)bitmapWidth, (float)bitmapHeight);
                    g.Dispose();
                }
                bitmap.Save(savePath + saveName + i.ToString("D3") + ".png", ImageFormat.Png);
                bitmap.Dispose();
                int percentage = (int)Math.Round(((double)i / (double)pageCount) * 100.0);
                if (OnProgressUpdate != null)
                {
                    OnProgressUpdate(Convert.ToInt32(percentage));
                }
                i++;
                Command.thisCommand.dialog.IncrementProgress();
                bitmap = null;
            }
            return pageCount;
        }

        private void setSheetParameters(double sWidth)
        {
            double swidth = sWidth;
            if (swidth == 3.0)
            {
                sheetSize = "24 x 36";
                Command.thisCommand.setSheetSize(sheetSize);
            }
            else if (swidth == 3.5)
            {
                sheetSize = "30 x 42";
                Command.thisCommand.setSheetSize(sheetSize);
            }
            else if (swidth == 4.0)
            {
                sheetSize = "36 x 48";
                Command.thisCommand.setSheetSize(sheetSize);
            }
            else
            {
                sheetSize = "void";
                Command.thisCommand.setSheetSize(sheetSize);
            }
        }

        private void setWordPath(string path)
        {
            wordDocPath = path;
        }

        private string getWordPath()
        {
            return wordDocPath;
        }

        private void removeExistingImages(string savePath)
        {
            DirectoryInfo X = new DirectoryInfo(savePath);
            FileInfo[] listOfFiles = X.GetFiles("*.png");
            string[] Collection = new string[listOfFiles.Length];

            foreach (FileInfo FI in listOfFiles)
            {
                string fileToDelete = savePath + FI.Name;
                File.Delete(fileToDelete);
            }
        }
    }
}
