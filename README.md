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
- "TableRepeatableSample" must be Repeating Section Content, and must be named with model class name below.
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
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingRepeatableTableStructureResult.jpg)
-----

### Writing Text to Repeatable Nested Block/Table Structure ###
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingNestedRepeatableStructure.jpg)

#### Define the Model Structure of Document ####
```csharp
// Reference MergeDocument
using MergeDocument.Model.BaseObject;
using MergeDocument.Model.Interface;

namespace MergeMocumentSampleCode.SampleModels
{
    // Template Class, must inheritance TemplateDataBase class
    public class NestedDocumentSample : TemplateDataBase
    {
        [BindField]
        public string OrderName { get; set; }

        // Block Repeater Property, must Add BlockRepeater attribute
        // Use collections as type of BlockRepeater Property
        [BlockRepeater]
        public List<ProductDescription> ProductDescriptions { get; set; }

        // Table Repeater Property, must Add Repeater attribute
        // Use collections as type of Repeater Property
        [Repeater]
        public List<OrderDetail> OrderDetails { get; set; }
    }

    // Nested Repeater
    public class ProductDescription : TemplateDataBase
    {
        [BindField]
        public string DescriptionContent { get; set; }
    }

    // Nested Repeater
    public class OrderDetail : TemplateDataBase
    {
        [BindField]
        public string UnitPrice { get; set; }

        [BindField]
        public string Quantity { get; set; }

        [BindField]
        public string Discount { get; set; }

        [BindField]
        public string Customer { get; set; }

        [BindField]
        public string ExtendedPrice { get; set; }
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
    public void NestedDocumentSample()
    {
        string output = Path.Combine(@"D:\", $"{DateTime.Now:yyyyMMddHHmmss}.docx");

        List<TemplateDataBase> data = new List<TemplateDataBase>
        {
            new NestedDocumentSample()
            {
               OrderName = "Pencils",
               ProductDescriptions = new List<ProductDescription>
               {
                   new ProductDescription
                   {
                       DescriptionContent ="Pencil DescriptionContent"
                   },
                   new ProductDescription
                   {
                       DescriptionContent ="Pencil Price"
                   },
                   new ProductDescription
                   {
                       DescriptionContent ="Pencil Supplier"
                   },
               },
               OrderDetails = new List<OrderDetail>
               {
                   new OrderDetail
                   {
                       UnitPrice = "10",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "10",
                       Quantity = "500",
                       Discount = "0",
                       Customer = "○○ Customer1",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "10",
                       Quantity = "300",
                       Discount = "0",
                       Customer = "○○ Customer2",
                       ExtendedPrice = "0",
                   },
               }
            },
            new NestedDocumentSample()
            {
               OrderName = "Notebooks",
               ProductDescriptions = new List<ProductDescription>
               {
                   new ProductDescription
                   {
                       DescriptionContent ="Notebooks DescriptionContent"
                   },
                   new ProductDescription
                   {
                       DescriptionContent ="Notebooks Price"
                   },
                   new ProductDescription
                   {
                       DescriptionContent ="Notebooks Supplier"
                   },
               },
               OrderDetails = new List<OrderDetail>
               {
                   new OrderDetail
                   {
                       UnitPrice = "40",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "40",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer1",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "40",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer2",
                       ExtendedPrice = "0",
                   },
               }
            },
            new NestedDocumentSample()
            {
               OrderName = "Erasers",
               ProductDescriptions = new List<ProductDescription>
               {
                   new ProductDescription
                   {
                       DescriptionContent ="Erasers DescriptionContent"
                   },
                   new ProductDescription
                   {
                       DescriptionContent ="Erasers Price"
                   },
                   new ProductDescription
                   {
                       DescriptionContent ="Erasers Supplier"
                   },
               },
               OrderDetails = new List<OrderDetail>
               {
                   new OrderDetail
                   {
                       UnitPrice = "12",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "12",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer1",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "12",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "○○ Customer2",
                       ExtendedPrice = "0",
                   },
               }
            },
        };

        var wordXmlBiz = new WordXMLBiz(data);
        wordXmlBiz.MergeFile(@"D:\NestedDocumentSample.docx", output);
        Console.WriteLine($@"File {output} generated successfully");
    }
}
```

#### Merge Result ####
