using System;
using TemplateEngine.Docx;
using System.IO;

namespace TemplateEngineDocxTest
{
    class Program
    {
        static void Main(string[] args)
        {
            File.Delete("output.docx");
            GenFile();
            var startInfo = new System.Diagnostics.ProcessStartInfo("output.docx") 
            {
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(startInfo);
        }

        static void GenFile()
        {
            var templateFileName = "template.docx";
            var tableContent = new TableContent("row");
            tableContent.AddRow(new FieldContent("subject", "數學"), new FieldContent("score", "90"));
            tableContent.AddRow(new FieldContent("subject", "物理"), new FieldContent("score", "80"));

            var listContent = new ListContent("Team Members List")
                    .AddItem(
                        new FieldContent("Name", "Eric"),
                        new FieldContent("Role", "Program Manager"))
                    .AddItem(
                        new FieldContent("Name", "Bob"),
                        new FieldContent("Role", "Developer"));

            var nestLiest = new ListContent("NestedList")
                    .AddItem(new ListItemContent("Role", "Program Manager")
                        .AddNestedItem(new FieldContent("Name", "Eric"))
                        .AddNestedItem(new FieldContent("Name", "Ann")))
                    .AddItem(new ListItemContent("Role", "Developer")
                        .AddNestedItem(new FieldContent("Name", "Bob"))
                        .AddNestedItem(new FieldContent("Name", "Richard")));

            var imageContent = new ImageContent("Image", File.ReadAllBytes("cat.jpg"));

            var valuesToFill = new Content(new FieldContent("name", "王大明"), new FieldContent("avg", "85"),tableContent, listContent, nestLiest, imageContent);
            using var file = new FileStream(templateFileName, FileMode.Open, FileAccess.Read);
            using var outputFileStream = new FileStream("output.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            file.CopyTo(outputFileStream);
            using var ouputDocument = new TemplateProcessor(outputFileStream).SetRemoveContentControls(true);
            ouputDocument.FillContent(valuesToFill);
            ouputDocument.SaveChanges();
        }
    }
}
