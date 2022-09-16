# NPOI.XWPFMapper

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

This is a library for mapping objects to Word document tables with the [NPOI](https://github.com/nissl-lab/npoi) library.

## Usage

> You can take a look at the `NPOI.XWPFMapper.ConsoleApp` project for a working example.

`NPOI.XWPFMapper` uses the interface `IXWPFMappable` and the attribute `XWPFPropertyAttribute` to define which classes should be mappable and which members should be nested or just strings.

```csharp
internal class ExampleClass : IXWPFMappable // Implement the IXWPFMappable interface to make a class mappable
{
    // Add the XWPFPropertyattribute to make a member mappable.
    [XWPFProperty("Color")] 
    public ExampleEnum Enum { get; set; }

    [XWPFProperty("Name")]
    public string Name { get; set; }

    // This child class also implements IXWPFMappable and will become a nested table.
    [XWPFProperty("Address")]
    public ExampleChildClass Address { get; set; }

    // Without the XWPFPropertyAttribute a member will be ignored by the mapping.
    public string IgnoredMember { get; set; }
}

internal class ExampleChildClass : IXWPFMappable
{
    [XWPFProperty("Street")]
    public string Address { get; set; }
    [XWPFProperty("Place")]
    public string City { get; set; }
    [XWPFProperty("Country")]
    public string CountryCode { get; set; }
}
```

In addition to this, there is also the `XWPFTableWrapper`, which is a class that manages an `XWPFTable`. It only accepts a single Type to be mapped (excluding its members of course).

```csharp
XWPFDocument document = new XWPFDocument();

ExampleClass exampleData = new ExampleClass()
{
    Enum = ExampleEnum.Red,
    Name = "This is a test",
    IgnoredMember = "This member will be ignored",
    Address = new ExampleChildClass()
    {
        Address = "Burgemeester Schönfeldplein",
        City = "Winschoten",
        CountryCode = "NL"
    }
};

// XWPFTableWrapper requires a Type argument and an XWPFDocument object to work
// In addition you can optionally set the direction of the table with an enum XWPFTableAlignment (default is Row)

XWPFTableWrapper<ExampleClass> wrapper = new XWPFTableWrapper<ExampleClass>(document, XWPFTableAlignment.Column);
wrapper.Insert(exampleData);

```

The table alignment refers to how the individual member values are inserted and not to how new records are added to the table.

`XWPFTableAlignment.Row`

| Color   | Red                       |
|---------|---------------------------|
| Name    | This is a test            |
| Address | Van Nelleweg 1, Rotterdam |

`XWPFTableAlignment.Column`

| Color | Name           | Address                   |
|-------|----------------|---------------------------|
| Red   | This is a test | Van Nelleweg 1, Rotterdam |

`XWPFTableWrapper`'s `XWPFTable` is public, so you can access it from `XWPFTableWrapper`.

## 🏳️‍⚧️ Trans Rights are human rights!
