# Cougar Config Support:
This `CougarConfigSupport` is a simple library for importing data from `Excel files (.xlsx)` and `JSON files (.json)` into a .NET application. It provides a list of classes and method to process data into desired part.


# Installation
To use the Config Support in your .NET project, follow these steps:

Ensure you have all this library installed in your project. You can install it via NuGet using the following command:
 - `Install-Package EPPlus`
 - `Install-Package Newtonsoft.Json`
 - `Install-Package CougarConfigSupport` Install latest package

# Usage
```
using CougarConfigSupport;

class Program
{
    private static OPCUAConversion yourClassInstance = new OPCUAConversion();
    private static PLCConversion yourClassInstance = new PLCConversion();

    static void Main(string[] args)
    {
        var responseTypeAccordingToMethodResponse = yourClassInstance.SomeMethod();
        //your code
    }
}
```

## OPCUAConversion Class:
The `OPCUAConversion` class provides a two methods
 - `FileToJSON` -- Convert Excel to JSON structure.
 - `JSONToFile` -- Convert JSON to Excel / JSON data with respect to the JSON

### FileToJSON Method
The `FileToJSON` method import Excel data and converting it to JSON. Here's how you can use it:

#### Parameters:
- `package`: An `ExcelPackage` object loaded with the Excel workbook.
- `SubProcess`: A list of sub-process names to identify and categorize specific nodes within the worksheet.
- `DataType`: A list of data types used to determine if a node's data type is basic or requires further processing.

#### Request:
CSharp:
```
using CougarConfigSupport;

class Program
{
    static void Main(string[] args)
    {
        OPCUAConversion yourClassInstance = new OPCUAConversion();
        using (var package = new ExcelPackage(new FileInfo("example.xlsx"))) // Specify the Excel file path
        {
            List<string> subProcess = new List<string> { "SubProcess1", "SubProcess2", "SubProcess3" };
            List<string> dataType = new List<string> { "DataType1", "DataType2", "DataType3" };

            OPCNodeConfig response = yourClassInstance.FileToJSON(package, subProcess, dataType);
            // Use the response as needed in your application
        }
    }
}
```
#### Response:
The return is an `OPCNodeConfig object` populated with data from the first worksheet in the provided Excel package.

Example Response:
```
{
  "description": "string",
  "version": "string",
  "opcNodes": [
    {
      "name": "string",
      "type": "string",
      "description": "string",
      "template": "string",
      "childTypes": [
        {
          "name": "string",
          "type": "string",
          "description": "string",
          "template": "string",
          "childTypes": []
        }
      ]
    }
  ]
}
```

### JSONToFile Method:
The `JSONToFile` method import JSON data and converting it to Excel / JSON files with respect to the JSON request. Here's how you can use it:

#### Parameters:
- `opcNode`: `OPCNode` type data.
- `pages`: List of `OPCPage` objects.

#### Example:
```
using CougarConfigSupport;

class Program
{
    public class OPCNode
    {
        public string name { get; set; }
        public string type { get; set; }
        public string template { get; set; }
        public string description { get; set; }
        public List<OPCNode> childTypes { get; set; }
        public List<string> ObjectLevels { get; set; } = new List<string>();
    }

    public class OPCPage
    {
        public string PageName { get; set; }
        public List<OPCColumnType> objects { get; set; } = new List<OPCColumnType>();
    }

    public class OPCColumnType
    {
        public string dataType { get; set; }
        public List<string> objectLevel { get; set; }
        public string tagName { get; set; }
        public string description { get; set; }
    }

    static void Main(string[] args)
    {
        OPCUAConversion yourClassInstance = new OPCUAConversion();

        // Prepare sample data
        OPCNode opcNode = new OPCNode
        {
            name = "RootNode",
            type = "ComplexType",
            description = "Root node description",
            template = "Template1",
            childTypes = new List<OPCNode>
            {
                new OPCNode
                {
                    name = "ChildNode1",
                    type = "SimpleType",
                    description = "Child node description",
                    template = "Template2",
                    childTypes = new List<OPCNode>()
                }
            }
        };

        List<OPCPage> pages = new List<OPCPage>
        {
            new OPCPage { PageName = "Template1" },
            new OPCPage { PageName = "Template2" }
        };

        // Call the JSONToFile method
        yourClassInstance.JSONToFile(opcNode, pages);

        // Use the 'pages' list as needed in your application
    }
}
```

#### Return Value:
No return value; the method will modify the `pages` list with the extracted data.

Example Response:
```
OPCNode opcNode = new OPCNode
{
    name = "RootNode",
    type = "ComplexType",
    description = "Root node description",
    template = "Template1",
    childTypes = new List<OPCNode>
    {
        new OPCNode
        {
            name = "ChildNode1",
            type = "SimpleType",
            description = "Child node 1 description",
            template = "Template2",
            childTypes = new List<OPCNode>()
        },
        new OPCNode
        {
            name = "ChildNode2",
            type = "ComplexType",
            description = "Child node 2 description",
            template = "Template3",
            childTypes = new List<OPCNode>
            {
                new OPCNode
                {
                    name = "GrandchildNode1",
                    type = "SimpleType",
                    description = "Grandchild node 1 description",
                    template = "Template4",
                    childTypes = new List<OPCNode>()
                }
            }
        }
    }
};

List<OPCPage> pages = new List<OPCPage>
{
    new OPCPage { PageName = "Template1" },
    new OPCPage { PageName = "Template2" },
    new OPCPage { PageName = "Template3" },
    new OPCPage { PageName = "Template4" }
};
```

## PLCConversion Class:
The `PLCConversion` class provides methods to read from and write to Excel files containing PLC configuration data. The methods include:
- `ExcelToJSON` -- Converts Excel data to a `PLCDictionaryDataType` object.
- `JSONToExcel` -- Converts a `PLCDictionaryDataType` object to an `Excel file`.

### ExcelToJSON Method
The `ExcelToJSON` method reads an Excel file and converts its content to a PLCDictionaryDataType object. Here's how you can use it:

#### Parameters:
- `filePath`: Specify the path of your Excel file location with name

#### Request:
CSharp:
```
using CougarConfigSupport;

class Program
{
    static void Main(string[] args)
    {
        PLCConversion yourClassInstance = new PLCConversion();
        string filePath = "example.xlsx"; // Specify the Excel file path

        Dictionary<string, List<string[]>> response = yourClassInstance.ConvertExcelToDictionary(string filePath);
        // Use the response as needed in your application
    }
}
```
#### Response:
The return type is `PLCDictionaryDataType`, populated with data from the `Excel file`.

Example Response:
```
{
  "name": "string",
  "version": "string",
  "comment": "string",
  "opcToPLCMappings": [
    {
      "plcType": "AllenBradley_Legacy",
      "plcChannel": "ABLegacyCLX_GW",
      "plcModel": "None",
      "plcAddress": "string",
      "refreshRate": 0,
      "manualRead": true,
      "connectionTimeout": 0,
      "transactionTimeout": 0,
      "connectionAttempts": 0,
      "enabled": true,
      "plcSettings1": "string",
      "plcSettings2": "string",
      "nodeMapping": [
        {
          "plcTag": "string",
          "plcTagType": "Byte",
          "accessType": "Write",
          "plcTagElement": 0,
          "modifyValueBy10": false,
          "modifyValueBy": 0,
          "description": "string",
          "template": "string",
          "opcNode": "string"
        }
      ]
    }
  ]
}
```

### JSONToExcel Method
The `JSONToExcel` method creates an Excel file from a PLCDictionaryDataType object. Here's how you can use it:

#### Parameters:
- `plcConfig`: The `PLCDictionaryDataType` object to be exported to Excel.
- `filePath`: The file path where the `Excel file` will be saved.

#### Request:
CSharp:
```
using CougarConfigSupport;

class Program
{
    static void Main(string[] args)
    {
        PLCConversion plcConversion = new PLCConversion();

        PLCDictionaryDataType plcConfig = new PLCDictionaryDataType
        {
            name = "ExampleName",
            version = "1.0",
            comment = "This is a sample PLC configuration.",
            opcToPLCMappings = new List<OpcToPLCMappings>
            {
                new OpcToPLCMappings
                {
                    plcType = "AllenBradley_Legacy",
                    plcChannel = "ABLegacyCLX_GW",
                    plcModel = "None",
                    plcAddress = "Address1",
                    refreshRate = 100,
                    manualRead = true,
                    connectionTimeout = 30,
                    transactionTimeout = 60,
                    connectionAttempts = 3,
                    enabled = true,
                    plcSettings1 = "Setting1",
                    plcSettings2 = "Setting2",
                    nodeMapping = new List<NodeMapping>
                    {
                        new NodeMapping
                        {
                            plcTag = "Tag1",
                            plcTagType = "Byte",
                            accessType = "Write",
                            plcTagElement = 1,
                            modifyValueBy10 = false,
                            modifyValueBy = 0,
                            description = "Description1",
                            template = "Template1",
                            opcNode = "OPCNode1"
                        }
                    }
                }
            }
        };
        string filePath = "output.xlsx"; // Specify the output Excel file path

        plcConversion.JSONToExcel(plcConfig, filePath);
        // The file will be saved at the specified file path
    }
}
```
#### Response:
`File` saved as `output.xlsx`


# Contributing
Contributions to the CougarConfigSupport library are welcome!


# License
This project is licensed under the MIT License


# Author Information
SIVA SHANMUGA VADIVEL