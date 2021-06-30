using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
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

        private void AddText(MainDocumentPart mainPart, string text)
        {
            Paragraph para = mainPart.Document.Body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));
        }

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