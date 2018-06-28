# MergeDocument.dll #
## Briefing Summary ##

MergeDocument.dll provides a simple and flexible structure that documents can be generated automatically through the combination of Microsoft World template and .NET Framework development custom model. Within some basic rules, you can do:

- Write document content with flexible custom model structure.
- Generate document content with repeatable template, including World Block and Table.
- Support .Net Framework development, and Microsoft World Developer template.

### Support ###
- Coding Environment: .Net Framework 4.6.1 Up
- Template File: Word 2013 Up

-----

## Example of usage ##

[Here](https://github.com/Itsower/MergeDocument "Here") you will find the example of source code and document templates.

### How to create template ###
- Open Microsoft Word and create new file.
- Click Developer tab.
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/wordDeveloperTag.jpg)

- If you cannot find the Developer tab on the Ribbon, [see](https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon "see").
- Create your own document template.
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/createDocumentTemplate.jpg)

-----

### Writing Text to Repeatable Block Structure ###
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

-----

### Writing Text to Repeatable Table Column Structure ### 
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingRepeatableTableStructure.jpg)
- "TableRepeatableSample" must be Repeating Section Content Control, and must be named with model class name below.
- Insert bindfields into the scope of "Table Repeatable Sample", the rule of bindfield is the same as block structure.

#### Define the Model Structure of Document ####
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

#### Merge Result ####
-----
