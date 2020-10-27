using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using eInvitationApp.Models;
using System.Data;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Microsoft.AspNetCore.Hosting;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.IO;
using GemBox.Document;

namespace eInvitationApp.Controllers
{
    public class HomeController : Controller
    {

        private readonly IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult Register(string id)
        {

            return View();
        }

        [HttpPost]
        public IActionResult Register(AttendeeModel obj)
        {

            dynamic file =  Download(obj.FirstName,obj.LastName);
            return View();
        }

        [HttpGet]
        public IActionResult Download(string firstName, string lastName)
        {

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var document = DocumentModel.Load("Resources/MOOLA2020.docx");
            document.Content.Replace("%Fullname%", firstName + " " + lastName, new CharacterFormat() { FontColor = Color.Black });
            byte[] fileContents;

            var options = GemBox.Document.SaveOptions.PdfDefault;

            // Save document to DOCX format in byte array.
            using (var stream = new MemoryStream())
            {
                document.Save(stream, options);

                fileContents = stream.ToArray();
            }

            // Stream document to browser in DOCX format.
            return File(fileContents, options.ContentType, "Resources/" + firstName + "_" + lastName + "_invite.pdf");
            //document.Save("Resources/"+ firstName+"_"+ lastName+"_invite.pdf");
        }

        //public string generatePDFdoc(Dictionary<string, string> Data, string name)
        //{
        //    try
        //    {
        //        //Loads an existing Word document //Policy DOcument
        //        //WordDocument wordDocument = new WordDocument(System.Web.Hosting.HostingEnvironment.MapPath("~/Resources/MOOLA2020.docx"), FormatType.Docx);
        //        //WordDocument wordDocument = new WordDocument("~/Resources/MOOLA2020.docx",FormatType.Docx);

        //        FileStream fileStreamPath = new FileStream(@"Resources/MOOLA2020.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        //        WordDocument wordDocument = new WordDocument(fileStreamPath, FormatType.Automatic);

        //        foreach (string key in Data.Keys)

        //        {
        //            #region Personalize PDF
        //            //Finds the first occurrence of a particular text in the document
        //            TextSelection textSelection = wordDocument.Find(key, false, true);
        //            if (textSelection != null)
        //            {
        //                //Gets the found text as single text range
        //                WTextRange textRange = textSelection.GetAsOneRange();
        //                //Modifies the text
        //                //textRange.Text = PolicyCertData[key]; //this replaces just one occurence
        //                wordDocument.Replace(key, Data[key], false, false);//replaces all occurences of the given text
        //            }
        //            #endregion
        //        }


        //        //Creates an instance of the DocToPDFConverter
        //        DocToPDFConverter converter = new DocToPDFConverter();
        //        //Converts Word document into PDF document
        //        PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
        //        //Saves the PDF file
        //        string Path = _hostingEnvironment.ContentRootPath + "Resources/" + name + "_invite.pdf";
        //        // new WordDocument(_hostingEnvironment.ContentRootPath + "/Resources/MOOLA2020.docx",FormatType.Docx);
        //        // Delete existing file first
        //        // Create the file again
        //        pdfDocument.Save(Path);
        //        //Closes the instance of document objects
        //        pdfDocument.Close(true);
        //        wordDocument.Close();
        //        return Path;
        //    }
        //    catch (Exception ex)
        //    {
        //        return null;
        //    }
        //}


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
