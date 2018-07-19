## Insert Image to Picture Content Control ##
### Document Template ###
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingNestedRepeatableStructureWithImage.jpg)
<br\>

### Define the Model Structure of Document ###
```csharp

// Reference MergeDocument
using MergeDocument.Model.BaseObject;
using MergeDocument.Model.Interface;

// Template Class, must inheritance TemplateDataBase class
public class InsertImageNestedSample : TemplateDataBase
{
    [BindField]
    public string ShipName { get; set; }

    // Block Repeater Property, must Add BlockRepeater attribute
    // Use collections as type of BlockRepeater Property
    [BlockRepeater]
    public List<ShipDescription> ShipDescriptions { get; set; }

    // Table Repeater Property, must Add Repeater attribute
    // Use collections as type of Repeater Property
    [TableRepeater]
    public List<ShippingDetail> ShippingDetails { get; set; }
}

// Nested Repeater
public class ShipDescription : TemplateDataBase
{
    [BindField]
    public string ShipDescriptionContent { get; set; }

    // Add Image Attribute
    [Image]
    public string ShipContentImage { get; set; }
}

// Nested Repeater
public class ShippingDetail : TemplateDataBase
{
    [BindField]
    public string ShippingPrice { get; set; }

    [BindField]
    public string ShippingQuantity { get; set; }

    [BindField]
    public string ShippingDiscount { get; set; }

    [BindField]
    public string ShipToCustomer { get; set; }

    [BindField]
    public string ShippingExtendedPrice { get; set; }

    // Add Image Attribute
    [Image]
    public string ShippingProductImage { get; set; }
}
```

#### Sample Code ####
```csharp
using MergeDocument.Biz;
using MergeDocument.Model.Interface;
using MergeMocumentSampleCode.SampleModels;
using MergeDocument.Common;

namespace MergeMocumentSampleCode
{
    public class DocumentBiz
    {
        public void InsertImageNestedDocumentSample()
        {
            string output = System.IO.Path.Combine(@"D:\", $"{DateTime.Now:yyyyMMddHHmmss}.docx");

            List<TemplateDataBase> data = new List<TemplateDataBase>
            {
                new InsertImageNestedSample()
                {
                   ShipName = "Cars",
                   ShipDescriptions = new List<ShipDescription>
                   {
                       new ShipDescription
                       {
                           ShipDescriptionContent ="Car Product DescriptionContent",
                           ShipContentImage = @""
                       },
                   },
                   ShippingDetails = new List<ShippingDetail>
                   {
                       new ShippingDetail
                       {
                           ShippingPrice = "10000000",
                           ShippingQuantity = "100",
                           ShippingDiscount = "0",
                           ShipToCustomer = "¡³¡³ Customer",
                           ShippingExtendedPrice = "0",
                           ShippingProductImage = @"D:\2LuRgfP.jpg"
                       },
                   }
                },
                new InsertImageNestedSample()
                {
                   ShipName = "4K TV",
                   ShipDescriptions = new List<ShipDescription>
                   {
                       new ShipDescription
                       {
                           ShipDescriptionContent ="4K TV Product DescriptionContent",
                           ShipContentImage = @"D:\2LuRgfP.jpg"
                       },
                       new ShipDescription
                       {
                           ShipDescriptionContent ="4K TV Product DescriptionContent",
                           ShipContentImage = @"D:\FIFA.png"
                       },
                   },
                   ShippingDetails = new List<ShippingDetail>
                   {
                       new ShippingDetail
                       {
                           ShippingPrice = "50000000",
                           ShippingQuantity = "100",
                           ShippingDiscount = "0",
                           ShipToCustomer = "¡³¡³ Customer",
                           ShippingExtendedPrice = "0",
                           ShippingProductImage = @"D:\FIFA.png"
                       },
                       new ShippingDetail
                       {
                           ShippingPrice = "10000000",
                           ShippingQuantity = "100",
                           ShippingDiscount = "0",
                           ShipToCustomer = "¡³¡³ Customer",
                           ShippingExtendedPrice = "0",
                           ShippingProductImage = @"D:\2LuRgfP.jpg"
                       },
                   }
                },
            };

            var wordXmlBiz = new WordXMLBiz(data);
            wordXmlBiz.MergeFile(@"D:\InsertImageNestedSample.docx", output);
            Console.WriteLine($@"File {output} generated successfully");
        }
    }
}

```

#### Merge Result ####
