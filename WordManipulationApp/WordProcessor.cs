using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using WordManipulationApp.Models;

namespace WordManipulationApp
{
    public class WordProcessor
    {


        public string ChartFile = AppDomain.CurrentDomain.BaseDirectory + "Files\\ChartComponent.docx";
        public string SignatureFile = AppDomain.CurrentDomain.BaseDirectory + "Files\\ClientSignatureComponent.docx";
        public string NameFile = AppDomain.CurrentDomain.BaseDirectory + "Files\\ClientNameComponent.docx";
        public string HeaderFile = AppDomain.CurrentDomain.BaseDirectory + "Files\\HeaderComponent.docx";
        public string PictureFile = AppDomain.CurrentDomain.BaseDirectory + "Files\\PictureComponent.docx";

        /// <summary>
        /// This method will add text and after adding text it will copy content from selected component word file.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="componentType"></param>
        /// <param name="OutputFile"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool ProcessWord(string text, ComponentEnum componentType, string OutputFile, int index)
        {
            using (WordprocessingDocument myDoc =
                WordprocessingDocument.Open(OutputFile, true))
            {
                string Chunk = "Chunk" + index;
                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                // Signature Code
                AddText(mainPart, text);
                CopyWord(GetFileName(componentType), mainPart, Chunk);

                mainPart.Document.Save();
            }
            return true;
        }

        /// <summary>
        /// This method will create an empty word document.
        /// </summary>
        /// <param name="path"></param>
        public void CreateEmptyDocument(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
            }
        }

        /// <summary>
        /// This method will copy content from passed file.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="mainPart"></param>
        /// <param name="altChunkId"></param>
        private void CopyWord(string file, MainDocumentPart mainPart, string altChunkId)
        {
            AlternativeFormatImportPart chunk =
                mainPart.AddAlternativeFormatImportPart(
                AlternativeFormatImportPartType.WordprocessingML, altChunkId);
            using (FileStream fileStream = File.Open(file, FileMode.Open))
                chunk.FeedData(fileStream);
            AltChunk altChunk = new AltChunk();
            altChunk.Id = altChunkId;

            mainPart.Document
                .Body
                .InsertAfter(altChunk, mainPart.Document.Body
                .Elements<Paragraph>().Last());
        }

        /// <summary>
        /// This method will add text in word document.
        /// </summary>
        /// <param name="mainPart"></param>
        /// <param name="text"></param>
        private void AddText(MainDocumentPart mainPart, string text)
        {
            Paragraph para = mainPart.Document.Body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));
        }

        /// <summary>
        /// This method will get file name based on passed component.
        /// </summary>
        /// <param name="componentType"></param>
        /// <returns></returns>
        private string GetFileName(ComponentEnum componentType)
        {
            var result = string.Empty;
            switch (componentType)
            {
                case ComponentEnum.Header:
                    result = HeaderFile;
                    break;
                case ComponentEnum.Name:
                    result = NameFile;
                    break;
                case ComponentEnum.Signature:
                    result = SignatureFile;
                    break;
                case ComponentEnum.Chart:
                    result = ChartFile;
                    break;
                case ComponentEnum.Picture:
                    result = PictureFile;
                    break;
            }
            return result;
        }
    }
}