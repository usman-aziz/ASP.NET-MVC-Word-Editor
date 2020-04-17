using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Word_Editor.Models;

namespace Word_Editor.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public FileResult Index(string editor)
        {
            try
            {
                // Create a unique file name
                string fileName = Guid.NewGuid() + ".docx";
                // Convert HTML text to byte array
                byte[] byteArray = Encoding.UTF8.GetBytes(editor.Contains("<html>") ? editor : "<html>" + editor + "</html>");
                // Generate Word document from the HTML
                MemoryStream stream = new MemoryStream(byteArray);
                Document Document = new Document(stream);
                // Create memory stream for the Word file
                var outputStream = new MemoryStream();
                Document.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0;
                // Return generated Word file
                return File(outputStream, System.Net.Mime.MediaTypeNames.Application.Rtf, fileName);
            }
            catch (Exception exp)
            {
                return null;
            }
        }
        [HttpPost]
        public ViewResult UploadFile(IFormFile file)
        {
            // Set file path
            var path = Path.Combine("wwwroot/uploads", file.FileName);
            using (var stream = new FileStream(path, FileMode.Create))
            {
                file.CopyTo(stream);
            }
            // Load Word document
            Document doc = new Document(path);
            var outStream = new MemoryStream();
            // Set HTML options
            HtmlSaveOptions opt = new HtmlSaveOptions();
            opt.ExportImagesAsBase64 = true;
            opt.ExportFontsAsBase64 = true; 
            // Convert Word document to HTML
            doc.Save(outStream, opt);
            // Read text from stream
            outStream.Position = 0;
            using(StreamReader reader = new StreamReader(outStream))
            {
                ViewBag.HtmlContent = reader.ReadToEnd();
            }
            return View("Index");
        }
    }
}
