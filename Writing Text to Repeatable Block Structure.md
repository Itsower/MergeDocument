### Writing Text to Repeatable Block Structure ###
#### Document Template ####
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingRepeatableStructure.jpg)

#### Define the Model Structure of Document ####
```csharp
// Reference MergeDocument
using MergeDocument.Model.BaseObject;
using MergeDocument.Model.Interface;

namespace MergeMocumentSampleCode.SampleModels
{
    // Template Class must inheritance TemplateDataBase class
    public class BlockRepeat : TemplateDataBase
    {
        // Bind Field Property, must Add BindField attribute
        [BindField]
        public string Text { get; set; }
    }
}
```

#### Sample Code ####
```csharp
// Reference MergeDocument
using MergeDocument.Biz;
using MergeDocument.Model.Interface;
using MergeMocumentSampleCode.SampleModels;
using MergeDocument.Common;

public class DocumentBiz
{
    public void BlockRepeatSample()
    {
        string output = Path.Combine(@"D:\", $"{DateTime.Now:yyyyMMddHHmmss}.docx");

        // New List of TemplateDataBase and add Template Class Object
        // Template Class must inheritance TemplateDataBase class
        List<TemplateDataBase> data = new List<TemplateDataBase>
        {
            new BlockRepeat { Text = "Write document content with flexible custom model structure." },
            new BlockRepeat { Text = "Generate document content with repeatable template, including World Block and Table." },
            new BlockRepeat { Text = "Support .Net Framework development, and Microsoft World Developer template." }
        };

        // Merge 
        var wordXmlBiz = new WordXMLBiz(data);
        wordXmlBiz.MergeFile(@"D:\BlockRepeatSample.docx", output);
        Console.WriteLine($@"File {output} generated successfully");
    }
}
```

#### Merge Result ####
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingRepeatableStructureResult.jpg)
