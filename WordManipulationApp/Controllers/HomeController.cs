using System;
using System.Collections.Generic;
using System.Web.Mvc;
using WordManipulationApp.Models;

namespace WordManipulationApp.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// Main method to return view.
        /// </summary>
        /// <returns></returns>
        public ActionResult Paragraph()
        {
            return View();
        }

        /// <summary>
        /// This method will parse text into multiple paragraphs.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        [HttpPost, ValidateInput(false)]
        public ActionResult Paragraph(ParagraphModel paragraph)
        {
            var indexOfClosingTag = paragraph.CombinedParagraph.IndexOf("</p>");
            paragraph.ParagraphList = new List<Component>();
            if (indexOfClosingTag == -1)
            {
                paragraph.ParagraphList.Add(new Component() { ParagraphText = paragraph.CombinedParagraph });
            }
            while (indexOfClosingTag > 0)
            {
                var componentText = paragraph.CombinedParagraph.Substring(0, indexOfClosingTag + 4);
                componentText = componentText.Replace("<p>", "");
                componentText = componentText.Replace("</p>", "");
                paragraph.ParagraphList.Add(new Component() { ParagraphText = componentText });
                paragraph.CombinedParagraph = paragraph.CombinedParagraph.Substring(indexOfClosingTag + 4);
                indexOfClosingTag = paragraph.CombinedParagraph.IndexOf("</p>");
            }
            TempData["paragraph"] = paragraph;
            return RedirectToAction("AssignComponent");
        }

        /// <summary>
        /// This method will return view to assign components to paragraphs.
        /// </summary>
        /// <returns></returns>
        public ActionResult AssignComponent()
        {
            var paragraph = (ParagraphModel)TempData.Peek("paragraph");

            return View(paragraph);
        }

        /// <summary>
        /// This method will assign components and download merged file.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        [HttpPost]
        public FileResult AssignComponent(ParagraphModel paragraph)
        {
            string OutputFile = AppDomain.CurrentDomain.BaseDirectory + "Files\\Output.docx";
            
            WordProcessor processor = new WordProcessor();
            processor.CreateEmptyDocument(OutputFile);

            var index = 0;
            foreach (var item in paragraph.ParagraphList)
            {
                processor.ProcessWord(item.ParagraphText, item.SelectedComponent, OutputFile, index);
                index++;
            }
            byte[] fileBytes = System.IO.File.ReadAllBytes(OutputFile);
            string fileName = "Output.docx";
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);            
        }

    }
}