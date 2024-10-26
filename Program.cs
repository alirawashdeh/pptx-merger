using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ShapeCrawler;
using A = DocumentFormat.OpenXml.Drawing;

namespace ConsoleApplication
{
    public class Program
    {    
        static readonly string outputFilename = "pptx/output.pptx";
        public static void Main(string[] args)
        {
            string[] pptxFilenames = GetPptxFilenames();
            if (pptxFilenames.Length == 0)
            {
                throw new ArgumentException("No pptx files found.");
            }

            string outputFilename = GetOutputFilename();
            Array.Copy(pptxFilenames, 0, pptxFilenames, 0, 1);
            AddSlidesFromFiles(pptxFilenames, outputFilename);
        }

        private static string[] GetPptxFilenames()
        {
            return Directory.GetFiles(Directory.GetCurrentDirectory(), "*.pptx", SearchOption.AllDirectories)
                .Where(f => !f.Contains("~$")).ToArray();
        }

        private static string GetOutputFilename()
        {
            return Path.Combine(Directory.GetCurrentDirectory(), "pptx", "output.pptx");
        }

        private static void AddSlidesFromFiles(string[] pptxFilenames, string outputFilename)
        {
            File.Copy(pptxFilenames[0], outputFilename, true);
            Console.Out.WriteLine("Copied file: " + pptxFilenames[0] + " to: " + outputFilename + " 👌");
            for (int i = 1; i < pptxFilenames.Length; i++)
            {
                AddSlides(pptxFilenames[i]);
                Console.Out.WriteLine("Copied slides from: " + pptxFilenames[i] + " to: " + outputFilename + " 👌");
            }
        }

    public static void AddSlides(string inputFilename)
    {
        var inputPres = new Presentation(inputFilename);
        var outputPres = new Presentation(outputFilename);
        
        foreach (ISlide slide in inputPres.Slides)
        {
            outputPres.Slides.Add(slide);
        }
        outputPres.Save();
    }

    }
}