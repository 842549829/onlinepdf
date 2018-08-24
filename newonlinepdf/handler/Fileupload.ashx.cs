using System;
using System.IO;
using System.Net;
using System.Web;
using Convert = Aspose.Pdf.Convert;

namespace newonlinepdf.handler
{
    /// <summary>
    /// Fileupload 的摘要说明
    /// </summary>
    public class Fileupload : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            var date = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            var showPath = $"/file/pdf/{date}.pdf";
            var pdfPath = HttpContext.Current.Server.MapPath("/file/pdf/") + date + ".pdf";
            HttpFileCollection files = HttpContext.Current.Request.Files;
            if (files.Count != 0)
            {
                var file = files[0];
                if (string.IsNullOrEmpty(file.FileName) == false)
                {
                    var filePath = HttpContext.Current.Server.MapPath("/file/") + Path.GetFileName(file.FileName);
                    file.SaveAs(HttpContext.Current.Server.MapPath("/file/") + Path.GetFileName(file.FileName));
                    var extension = Path.GetExtension(file.FileName).ToLower();
                    if (extension == ".doc")
                    {
                        Convert.ConvertWordToPdf(filePath, pdfPath);
                    }
                    else if (extension == ".pptx" || extension == ".ppt")
                    {
                        Convert.ConvertPptToPdf(filePath, pdfPath);
                    }
                    else if (extension == ".xls" || extension == ".xlsx")
                    {
                        Convert.ConvertExcelToPdf(filePath, pdfPath);
                    }
                    else
                    {
                        showPath = "fail";
                    }
                }
            }
            context.Response.Write(showPath);
            context.Response.End();
        }

        public bool IsReusable => false;
    }
}