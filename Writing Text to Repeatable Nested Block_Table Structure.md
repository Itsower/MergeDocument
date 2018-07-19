## Writing Text to Repeatable Nested Block/Table Structure ##
### Document Template ###
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingNestedRepeatableStructure.jpg)

### Define the Model Structure of Document ###
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

### Sample Code ###
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
                       Customer = "¡³¡³ Customer",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "10",
                       Quantity = "500",
                       Discount = "0",
                       Customer = "¡³¡³ Customer1",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "10",
                       Quantity = "300",
                       Discount = "0",
                       Customer = "¡³¡³ Customer2",
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
                       Customer = "¡³¡³ Customer",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "40",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "¡³¡³ Customer1",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "40",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "¡³¡³ Customer2",
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
                       Customer = "¡³¡³ Customer",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "12",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "¡³¡³ Customer1",
                       ExtendedPrice = "0",
                   },
                   new OrderDetail
                   {
                       UnitPrice = "12",
                       Quantity = "100",
                       Discount = "0",
                       Customer = "¡³¡³ Customer2",
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

### Merge Result ###
![Alt text](https://github.com/Itsower/MergeDocument/blob/master/writingNestedRepeatableStructureResult.jpg)