## Writing Text to Repeatable Table Column Structure ##
### Document Template ###
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingRepeatableTableStructure.jpg)
- "TableRepeatableSample" must be Repeating Section Content, and must be named with model class name below.
- Insert bindfields into the scope of "Table Repeatable Sample", the rule of bindfield is the same as block structure.

### Define the Model Structure of Document ###
```csharp
// Reference MergeDocument
using MergeDocument.Model.BaseObject;
using MergeDocument.Model.Interface;

namespace MergeMocumentSampleCode.SampleModels
{
    // Template Class must inheritance TemplateDataBase class
    public class TableRepeatSample : TemplateDataBase
    {
        [BindField]
        public string Name { get; set; }

        [BindField]
        public string InputTypeName { get; set; }

        [BindField]
        public string Descriptions { get; set; }
    }
}
```

### Sample Code ###
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

            List<TemplateDataBase> data = new List<TemplateDataBase>
            {
                new TableRepeatSample
                {
                    Name = "Login Button",
                    InputTypeName = "Button",
                    Descriptions = "Login Button Descriptions"
                },
                new TableRepeatSample
                {
                    Name = "Cancel Button",
                    InputTypeName = "Button",
                    Descriptions = "Cancel Button Descriptions"
                },
                new TableRepeatSample
                {
                    Name = "Account Input",
                    InputTypeName = "TextBox",
                    Descriptions = "Account Input Descriptions"
                },
                new TableRepeatSample
                {
                    Name = "Password Input",
                    InputTypeName = "TextBox",
                    Descriptions = "Password Input Descriptions"
                },
            };

        // set WordXNLBiz--> initialRepeaterType = SystemEnum.RepeaterType.TableRow
        // initialRepeaterType default is SystemEnum.RepeaterType.Block
        var wordXmlBiz = new WordXMLBiz(data, SystemEnum.RepeaterType.TableRow);
        
        wordXmlBiz.MergeFile(@"D:\TableRepeatSample.docx", output);
        Console.WriteLine($@"File {output} generated successfully");
    }
}
```

### Merge Result ###
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingRepeatableTableStructureResult.jpg)