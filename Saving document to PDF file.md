## Saving document to PDF file ##
#### Sample Code ####
```csharp
// using Merge Document Component
using MergeDocument.Biz;
using MergeDocument.Model.Interface;
using MergeMocumentSampleCode.SampleModels;
using MergeDocument.Common;

namespace MergeMocumentSampleCode
{
    public class DocumentBiz
    {
        /// <summary>
        /// Word Block Repeat
        /// </summary>
        public void BlockRepeatSample()
        {
            // Output File Name
            string output = System.IO.Path.Combine(@"D:\", $"{DateTime.Now:yyyyMMddHHmmss}.docx");

            List<TemplateDataBase> data = new List<TemplateDataBase>
            {
                new BlockRepeat { Text = "Write document content with flexible custom model structure." },
                new BlockRepeat { Text = "Generate document content with repeatable template, including World Block and Table." },
                new BlockRepeat { Text = "Support .Net Framework development, and Microsoft World Developer template." }
            };

            var wordXmlBiz = new WordXMLBiz(data);

            // Set generatePDF is true (3rd arguement generatePDF = true of MergeFile method, default of this is false)
            wordXmlBiz.MergeFile(@"D:\BlockRepeatSample.docx", output, true);
            
            Console.WriteLine($@"File {output} generated successfully");
        }
    }
}
```

## Generate Result ##
