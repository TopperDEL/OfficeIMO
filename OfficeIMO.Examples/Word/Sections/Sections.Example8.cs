using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {
        internal static void Example_SectionsWithHeadersAndParagraphs(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with sections and headers / footers and paragraphs ");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections and headers footers and paragraphs.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                document.Sections[0].Header.First.AddParagraph().SetText("Test Section 0 - First Header");
                document.Sections[0].Header.Default.AddParagraph().SetText("Test Section 0 - Header");
                document.Sections[0].Header.Even.AddParagraph().SetText("Test Section 0 - Even");

                document.AddPageBreak();


                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();
                section1.Header.Default.AddParagraph().SetText("Test Section 1 - Header");
                section1.DifferentFirstPage = true;
                section1.Header.First.AddParagraph().SetText("Test Section 1 - First Header");


                Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Header.Default.Paragraphs[0].Text);
                Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Header.Default.Paragraphs[0].Text);

                Console.WriteLine("Section 0 - Header Paragraphs: " + document.Sections[0].Header.Default.Paragraphs.Count);
                Console.WriteLine("Section 1 - Header Paragraphs: " + document.Sections[1].Header.Default.Paragraphs.Count);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                //Console.WriteLine("-----");
                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 1: " + document.Sections[2].Paragraphs[1].Text);
                //Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Paragraphs[0].Text);
                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 1 - Text 0: " + document.Sections[1].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 2 - Text 0: " + document.Sections[2].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("Section 3 - Text 0: " + document.Sections[3].Header.Default.Paragraphs[0].Text);
                //Console.WriteLine("-----");
                //document.Sections[1].Header.Default.AddParagraph().SetText("Test Section 1 - Header-Par1");
                //Console.WriteLine("Section 1 - Text 1: " + document.Sections[1].Header.Default.Paragraphs[1].Text);
                document.Save(openWord);
            }
        }



    }
}
