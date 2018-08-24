using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using Aspose.Words.Saving;
using SaveFormat = Aspose.Words.SaveFormat;

namespace Aspose.Pdf
{
    /// <summary>
    /// 转化
    /// </summary>
    public class Convert
    {
        #region Word
        public static void ConvertWordToPdf(string wordInputPath, string imageOutputPath)
        {
            Document document = new Document(wordInputPath);
            document.Save(imageOutputPath, SaveFormat.Pdf);
        }

        public static void ConvertWordToHtml(string wordInputPath, string imageOutputPath)
        {
            Document document = new Document(wordInputPath);
            document.Save(imageOutputPath, SaveFormat.Html);
        }

        public static void ConvertWordToImage(string wordInputPath, string imageOutputPath, string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, float resolution)
        {
            try
            {
                Document doc = new Document(wordInputPath);
                if (imageOutputPath.Trim().Length == 0)
                {
                    imageOutputPath = Path.GetDirectoryName(wordInputPath);
                }
                if (!Directory.Exists(imageOutputPath))
                {
                    Directory.CreateDirectory(imageOutputPath ?? throw new ArgumentNullException(nameof(imageOutputPath)));
                }
                if (imageName.Trim().Length == 0)
                {
                    imageName = Path.GetFileNameWithoutExtension(wordInputPath);
                }
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.PageCount || endPageNum <= 0)
                {
                    endPageNum = doc.PageCount;
                }
                if (startPageNum > endPageNum)
                {
                    startPageNum = endPageNum;
                    endPageNum = startPageNum;
                }
                if (imageFormat == null)
                {
                    imageFormat = ImageFormat.Png;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }

                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imageFormat))
                {
                    Resolution = resolution
                };

                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    imageSaveOptions.PageIndex = i - 1;
                    doc.Save(Path.Combine(imageOutputPath, imageName ?? throw new ArgumentNullException(nameof(imageName))) + "_" + i + "." + imageFormat, imageSaveOptions);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static SaveFormat GetSaveFormat(ImageFormat imageFormat)
        {
            SaveFormat sf;
            if (imageFormat.Equals(ImageFormat.Png))
            {
                sf = SaveFormat.Png;
            }
            else if (imageFormat.Equals(ImageFormat.Jpeg))
            {
                sf = SaveFormat.Jpeg;
            }
            else if (imageFormat.Equals(ImageFormat.Tiff))
            {
                sf = SaveFormat.Tiff;
            }
            else if (imageFormat.Equals(ImageFormat.Bmp))
            {
                sf = SaveFormat.Bmp;
            }
            else
            {
                sf = SaveFormat.Unknown;
            }
            return sf;
        } 
        #endregion

        public static void ConvertExcelToPdf(string srcDocPath, string dstPdfPath)
        {
            Workbook wb = new Workbook(srcDocPath);
            wb.Save(dstPdfPath);
        }

        public static void ConvertPptToPdf(string srcDocPath, string dstPdfPath)
        {
            Presentation ppt = new Presentation(srcDocPath);
            ppt.Save(dstPdfPath, Slides.Export.SaveFormat.Pdf);
        }
    }
}