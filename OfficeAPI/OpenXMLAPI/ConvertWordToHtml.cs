using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeAPI.OpenXMLAPI
{
    public class ConvertWordToHtml:IButtonClick
    {
        
        public static void WordToHtml()
        {
            byte[] byteArray = File.ReadAllBytes(@"D:\OfficeDev\Word\201701\Test.docx");
            string css = @"
        p.PtNormal
            {margin-bottom:10.0pt;
            font-size:11.0pt;
            font-family:""Times"";}
        span.PtDefaultParagraphFont
            {margin-top:24.0pt;
            font-size:14.0pt;
            font-family:""Helvetica"";
            color:yellow;}
        h1.PtHeading1
            {margin-top:24.0pt;
            font-size:14.0pt;
            font-family:""Helvetica"";
            color:blue;}
        h2.PtHeading2
            {margin-top:10.0pt;
            font-size:13.0pt;
            font-family:""Helvetica"";
            color:blue;}";
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    int imageCounter = 0;
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "My Page Title",
                        
                        CssClassPrefix = "Pt",
                        AdditionalCss= css,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(@"D:\OfficeDev\Word\201701\img");
                            if (!localDirInfo.Exists)
                            {
                                localDirInfo.Create();
                            }
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }
                            if (imageFormat == null)
                                return null;

                            string imageFileName = @"D:\OfficeDev\Word\201701\image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                    File.WriteAllText(@"D:\OfficeDev\Word\201701\WordHtml.html", html.ToStringNewLineOnAttributes());
                };
            }
        }

        public void MessageBox(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }
    }
}
