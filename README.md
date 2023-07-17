# What does it do?

This library aims to handle the import of data from an excel spreadsheet.
In the future this library can be adjusted to work with other file formats.

# How to use it?

The library reads the records from the import file and maps them to a complex object, 
so the first implementation step is to define the object type and map it.

Where: the first element on mapping is the column header of the import file and the second is the property on the entity.

```csharp
    public class Product : Importable<Product>
    {
        private readonly IImportFileMapping<Product> _headerMapping = new ImportFileMapping<Product>()
            .MapRequired("Id", p => p.Id)
            .MapRequired("Product Name", p => p.Name)
            .MapOptional("Bar Code", p => p.Sku);

        public override IImportFileMapping<EntityValidImportableTest> HeaderMapping => _headerMapping;

        public Guid Id { get; set; }
        public string Name { get; set; }
        public string Sku { get; set; }
    }
```

In the mapping model it is possible to define which columns are required and optional using the methods **MapRequired** and **MapOptional**.

Once the mapping is done, just instantiate the read object.

```csharp
    public Product ReadFirst(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        return importReader.ReadLine();
    }
```

```csharp
    public IEnumerable<Product> ReadAll(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        return importReader.ReadAllLines();
    }
```

# Import Log

The library allows writing logs to the import file. 
Internally, the library generates an output file containing exactly the same layout and records that are pointed out in error. 
Additionally a column is added to indicate the errors found.

```csharp
    public void WriteLogFile(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        var product = importReader.ReadLine();

        var errors = new List<ValidationFailure>()
        {
            new ValidationFailure(null, "Invalid Name"),
            new ValidationFailure(null, "Invalid BarCode")
        };

        importReader.SetErrorsInLogFile(errors);
        importReader.SaveLog("C:/Log.xlsx")
    }
```

In the above implementation, the last line read is the one that will be generated in the log file.
If it is necessary to indicate a specific line, an overload method can be used.

```csharp
    public void WriteLogFile(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        var product = importReader.ReadLine();

        var errors = new List<ValidationFailure>()
        {
            new ValidationFailure(null, "Invalid Name"),
            new ValidationFailure(null, "Invalid BarCode")
        };

        importReader.SetErrorsInLogFile(product.SourceImportRowIndex, errors);
        importReader.SaveLog("C:/Log.xlsx")
    }
```

#### Example:

#### Import File

| Id  | Product Name | Bar Code |
| --- | ------------ | -------- |
| 123 | Bean         | 0987654  |

#### Log File

| Id  | Product Name | Bar Code | Errors |
| --- | ------------ | -------- | ------ |
| 123 | Bean         | 0987654  | Invalid Name \| Invalid BarCode |

# Methods

#### GetLogAsMemoryStream

Used to generate a memory stream from the import file.

```csharp
    public MemoryStream GetLogFileMemoryStream(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        importReader.ReadLine();

        var errors = new List<ValidationFailure>()
        {
            new ValidationFailure(null, "Invalid Name"),
        };

        importReader.SetErrorsInLogFile(product.SourceImportRowIndex, errors);
        return importReader.GetLogAsMemoryStream()
    }
```

#### Validate

Used to validate that the import file is valid.

```csharp
    public IEnumerable<ValidationFailure> GetValidationErrors(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        
        var validationResult = importReader.Validate();

        if (!validationResult.IsValid)
            return validationResult.Errors;

        return null;
    }
```

For a file to be considered valid, the following conditions must be true:

- The file must contain at least one import record.
- The file must contain all columns of the required type that were mapped in the output entity.
- The file cannot have duplicate columns.

#### CountTotalValidRegisters

Returns the number of records where at least one of the mapped columns has data. Ignore blank lines.

```csharp
    public bool FileHasValidData(MemoryStream importFileMemoryStream)
    {
        using var importReader = new ExcelFileImportReader<Product>(importFileMemoryStream);
        return importReader.CountTotalValidRegisters() > 0;
    }
```