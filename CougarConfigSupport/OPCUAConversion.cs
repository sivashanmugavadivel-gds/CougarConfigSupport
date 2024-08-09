using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;

namespace CougarConfigSupport
{
    public class PackageTagsAttribute : Attribute
    {
        public string[] Tags { get; }

        public PackageTagsAttribute(params string[] tags)
        {
            Tags = tags;
        }
    }
    public class OPCUAConversion
    {
        ///////////////////////////////////////File To JSON//////////////////////////////////////////////
        
        private string description;
        private string subProcessType;
        /// <summary>
        /// Processes the first worksheet in the given Excel package and converts it into an OPCNodeConfig.
        /// This model represents the structured data extracted from the Excel sheet, including descriptions,
        /// version information, and hierarchical node structures that are typical in OPC configurations.
        /// </summary>
        /// <param name="package">An ExcelPackage object loaded with the Excel workbook. The workbook should contain
        /// at least one worksheet that conforms to the expected layout for processing.</param>
        /// <param name="SubProcess">A list of sub process names to identify and categorize specific nodes within the worksheet.</param>
        /// <param name="DataType">A list of data types used to determine if a node's data type is basic or requires further processing.</param>
        /// <returns>
        /// An OPCNodeConfig object populated with data from the first worksheet in the provided Excel package.
        /// If the worksheet does not exist or does not match expected formats, the method handles errors gracefully and
        /// provides an appropriate empty or default-initialized OPCNodeConfig.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the "OPC Config Model" sheet is not found or if the sheet does not contain proper data to extract.
        /// </exception>
        public OPCNodeConfig FileToJSON(ExcelPackage package, List<string> SubProcess, List<string> DataType)
        {
            try
            {
                var mainSheet = package.Workbook.Worksheets["OPC Config Model"];
                var sheetNames = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
                var similarSheets = FindSimilarSheetNames("OPC Config Model", sheetNames);
                string errorMessage = string.Empty;

                #region Excel Sheet Exception
                if (mainSheet == null)
                {
                    errorMessage = $"\"OPC Config Model\" sheet not found! Similar sheets found: {string.Join(", ", similarSheets)}";
                    Console.WriteLine(errorMessage);
                    throw new Exception(errorMessage);  // Throw an exception indicating the sheet could not be found.
                }
                if (mainSheet.Dimension == null || mainSheet.Dimension.End.Row < 2)
                {
                    errorMessage = "Sheet \"OPC Config Model\" not have proper data to extract";
                    Console.WriteLine(errorMessage);
                    throw new Exception(errorMessage);  // Throw an exception indicating the sheet could not be found.
                }
                #endregion Excel Sheet Exception

                // Calculate all formulas in the worksheet
                mainSheet.Calculate();
                OPCNodeConfig configModel = new OPCNodeConfig
                {
                    description = mainSheet.Cells[1, 2].Text,  // OPC Config Model sheet description value cell B1
                    version = mainSheet.Cells[2, 2].Text,  // OPC Config Model sheet description value cell B2
                    opcNodes = new List<OPCNode>()
                };

                Dictionary<string, OPCNode> pathToNodeMap = new Dictionary<string, OPCNode>();
                string currentPath = "";
                subProcessType = "";

                for (int row = 4; row <= mainSheet.Dimension.End.Row; row++)
                {
                    OPCNode lastNode = null;
                    currentPath = "";  // Reset path at the start of each row

                    for (int col = 1; col <= 10; col++)  // Assuming columns 1 to 10 represent hierarchical levels
                    {
                        if (!string.IsNullOrEmpty(mainSheet.Cells[row, col].Text))
                        {
                            currentPath += "/" + mainSheet.Cells[row, col].Text;  // Build the unique path

                            if (!pathToNodeMap.TryGetValue(currentPath, out lastNode))
                            {
                                subProcessType = SubProcess.Contains(mainSheet.Cells[row, col].Text) ? mainSheet.Cells[row, col].Text : subProcessType;  // Check and update the sub process type e.g: WC1, WC2, COM..
                                // Create new node if it doesn't exist
                                lastNode = new OPCNode
                                {
                                    name = mainSheet.Cells[row, col].Text,
                                    type = "Empty",
                                    childTypes = new List<OPCNode>()
                                };
                                pathToNodeMap[currentPath] = lastNode;

                                // Link the node to its parent, if any
                                if (col > 1)
                                {
                                    string parentPath = currentPath.Substring(0, currentPath.LastIndexOf('/'));
                                    if (pathToNodeMap.TryGetValue(parentPath, out OPCNode parentNode))
                                    {
                                        parentNode.childTypes.Add(lastNode);
                                    }
                                }
                                else
                                {
                                    configModel.opcNodes.Add(lastNode);  // Add root node directly under OPCNodes
                                }
                            }
                        }
                    }

                    // Additional processing for properties in columns 15 and 16 if lastNode is identified
                    if (lastNode != null || !string.IsNullOrEmpty(mainSheet.Cells[row, 15].Text))
                    {
                        bool isBasicDataType = IsBasicDataType(mainSheet.Cells[row, 16].Text, DataType);
                        description = isBasicDataType ? mainSheet.Cells[row, 17].Text : CheckForPatternMatch(mainSheet.Cells[row, 17].Text) ? "" : mainSheet.Cells[row, 17].Text;  // Update description with normal or specific type

                        subProcessType = SubProcess.Contains(mainSheet.Cells[row, 15].Text) ? mainSheet.Cells[row, 15].Text : subProcessType;  // Check and update the sub process type

                        OPCNode detailNode = new OPCNode
                        {
                            name = mainSheet.Cells[row, 15].Text,
                            type = isBasicDataType ? mainSheet.Cells[row, 16].Text : "Empty",
                            template = "Empty",  // Main sheet does not have template
                            description = description,
                            childTypes = new List<OPCNode>()
                        };


                        if (!isBasicDataType && !string.IsNullOrEmpty(mainSheet.Cells[row, 16].Text))  // Process additional child types
                        {
                            detailNode.childTypes = ProcessChildType(package, mainSheet.Cells[row, 16].Text, subProcessType, SubProcess, DataType, CheckForPatternMatch(mainSheet.Cells[row, 17].Text) ? mainSheet.Cells[row, 17].Text : "");
                        }

                        if (lastNode == null)
                        {
                            if (!string.IsNullOrEmpty(mainSheet.Cells[row, 17].Text) && !IsBasicDataType(mainSheet.Cells[row, 17].Text, DataType))
                            {
                                // Handling column 17 as a new root node if it is not a basic data type
                                CougarConfigSupport.OPCNode column17Node = new CougarConfigSupport.OPCNode
                                {
                                    name = mainSheet.Cells[row, 17].Text,
                                    type = "Complex",
                                    template = "Empty",
                                    description = CheckForPatternMatch(mainSheet.Cells[row, 17].Text) ? "" : mainSheet.Cells[row, 17].Text,
                                    childTypes = ProcessChildType(package, mainSheet.Cells[row, 17].Text, subProcessType, SubProcess, DataType, "")
                                };
                                configModel.opcNodes.Add(column17Node);
                            }
                            else
                            {
                                // Assume detailNode from column 15 is a root node if lastNode is null
                                configModel.opcNodes.Add(detailNode);
                            }
                        }
                        else
                        {
                            // Normal case, add detailNode under lastNode
                            lastNode.childTypes.Add(detailNode);
                        }
                    }
                }

                return configModel;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Processes a specified worksheet within the given Excel package to extract child types based on a specified sub process type and an optional description pattern.
        /// This method is typically used to parse and transform specific sections of an Excel sheet into structured data of child types,
        /// often used in configurations or hierarchical data structures.
        /// </summary>
        /// <param name="package">An ExcelPackage object containing the workbook with the target worksheet. This package should be pre-loaded with the workbook that contains the sheet of interest.</param>
        /// <param name="sheetName">The name of the worksheet within the Excel workbook that will be processed. This sheet should contain the data relevant to the child types to be extracted.</param>
        /// <param name="subProcessType">A string specifying the sub process type that helps in filtering or categorizing the data during processing.</param>
        /// <param name="SubProcess">A list of subprocess names to identify and categorize specific nodes within the worksheet.</param>
        /// <param name="DataType">A list of data types used to determine if a node's data type is basic or requires further processing.</param>
        /// <param name="descPattern">An optional regex pattern string used for matching descriptions in the sheet for further refined data extraction. If null, no pattern matching is applied.</param>
        /// <returns>A list of OPCNode objects each representing a structured form of the data rows extracted based on the sub process type and description pattern.</returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the specified worksheet is not found or if the sheet does not contain proper data to extract.
        /// </exception>
        private List<OPCNode> ProcessChildType(ExcelPackage package, string sheetName, string subProcessType, List<string> SubProcess, List<string> DataType, string descPattern = null)
        {
            try
            {
                var sheet = package.Workbook.Worksheets[sheetName];
                var sheetNames = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
                var similarSheets = FindSimilarSheetNames(sheetName, sheetNames);
                List<OPCNode> childTypes = new List<OPCNode>();
                string errorMessage = string.Empty;

                #region Sheet Exception
                if (sheet == null)
                {
                    errorMessage = $"\"{sheetName}\" sheet not found! Similar sheets found: {string.Join(", ", similarSheets)}";
                    Console.WriteLine($"\"{sheetName}\" sheet not found! Similar sheets found: {string.Join(", ", similarSheets)}");
                    throw new Exception(errorMessage);  // Throw an exception indicating the sheet could not be found.
                }
                if (sheet.Dimension == null || sheet.Dimension.End.Row < 2)
                {
                    errorMessage = $"Sheet \"{sheetName}\" not have proper data to extract";
                    Console.WriteLine(errorMessage);
                    throw new Exception(errorMessage);  // Throw an exception indicating the sheet could not be found.
                }
                #endregion Sheet Exception

                // Calculate all formulas in the worksheet
                sheet.Calculate();
                Dictionary<string, OPCNode> pathToNodeMap = new Dictionary<string, OPCNode>();
                string currentPath = "";

                for (int row = 2; row <= sheet.Dimension.End.Row; row++)  // Assuming Child sheet row starts from 2
                {
                    OPCNode lastNode = null;
                    currentPath = "";  // Reset path at the start of each row

                    for (int col = 1; col <= 10; col++)  // Assuming columns 1 to 10 represent hierarchical levels
                    {
                        if (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
                        {
                            currentPath += "/" + sheet.Cells[row, col].Text;  // Build the unique path

                            if (!pathToNodeMap.TryGetValue(currentPath, out lastNode))
                            {
                                subProcessType = SubProcess.Contains(sheet.Cells[row, col].Text) ? sheet.Cells[row, col].Text : subProcessType;  // Check and update the sub process type e.g: WC1, WC2, COM..
                                // Create new node if it doesn't exist
                                lastNode = new OPCNode
                                {
                                    name = sheet.Cells[row, col].Text,
                                    type = "Empty",
                                    childTypes = new List<OPCNode>()
                                };
                                pathToNodeMap[currentPath] = lastNode;

                                // Link the node to its parent, if any
                                if (col > 1)
                                {
                                    string parentPath = currentPath.Substring(0, currentPath.LastIndexOf('/'));
                                    if (pathToNodeMap.TryGetValue(parentPath, out OPCNode parentNode))
                                    {
                                        parentNode.childTypes.Add(lastNode);
                                    }
                                }
                                else
                                {
                                    childTypes.Add(lastNode);  // Add root node directly under OPCNodes
                                }
                            }
                        }
                    }

                    // Additional processing for properties in columns 15 and 16 if lastNode is identified
                    if (!string.IsNullOrEmpty(sheet.Cells[row, 15].Text))
                    {
                        bool isBasicDataType = IsBasicDataType(sheet.Cells[row, 16].Text, DataType);
                        description = isBasicDataType ? sheet.Cells[row, 17].Text : CheckForPatternMatch(sheet.Cells[row, 17].Text) ? "" : sheet.Cells[row, 17].Text;  // Update description with normal or specific type
                        if (!string.IsNullOrEmpty(descPattern))  // Update specific type description with sub process xx and yy
                        {
                            description = ReplaceValues(descPattern, subProcessType, sheet.Cells[row, 17].Text);
                        }

                        subProcessType = SubProcess.Contains(sheet.Cells[row, 15].Text) ? sheet.Cells[row, 15].Text : subProcessType;  // Check and update the sub process type

                        OPCNode detailNode = new OPCNode
                        {
                            name = sheet.Cells[row, 15].Text,
                            type = isBasicDataType ? sheet.Cells[row, 16].Text : "Empty",
                            template = sheetName,
                            description = description,
                            childTypes = new List<OPCNode>()
                        };

                        if (!isBasicDataType && !string.IsNullOrEmpty(sheet.Cells[row, 16].Text))  // Process additional child types
                        {
                            detailNode.childTypes = ProcessChildType(package, sheet.Cells[row, 16].Text, subProcessType, SubProcess, DataType, CheckForPatternMatch(sheet.Cells[row, 17].Text) ? sheet.Cells[row, 17].Text : "");
                        }
                        if (lastNode == null)
                            childTypes.Add(detailNode);
                        else
                            lastNode.childTypes.Add(detailNode);
                    }
                }

                return childTypes;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Identifies and returns a list of sheet names that start with the same prefix as the specified desired sheet name.
        /// This method is useful for finding sheets with similar naming conventions, particularly when the exact name might vary slightly but start similarly.
        /// </summary>
        /// <param name="desiredSheetName">The sheet name to compare against other sheet names in the workbook. The comparison uses only the first few characters (up to 3).</param>
        /// <param name="sheetNames">A list of all sheet names available in the workbook. This list is searched to find matches based on the prefix.</param>
        /// <returns>
        /// A list of sheet names that start with the same prefix as the specified desired sheet name.
        /// The prefix length used for comparison is the lesser of 3 or the length of the desired sheet name.
        /// This comparison is case-insensitive to ensure a broad match across different naming cases.
        /// </returns>
        private static List<string> FindSimilarSheetNames(string desiredSheetName, List<string> sheetNames)
        {
            var maxLength = Math.Min(desiredSheetName.Length, 3); // Look for names that match the first 3 characters
            var prefix = desiredSheetName.Substring(0, maxLength);

            return sheetNames.Where(name => name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToList();
        }
        /// <summary>
        /// Determines whether the specified data type is considered a basic data type.
        /// A basic data type is one that is predefined and recognized as a standard or primitive type,
        /// which is typically used for validation and processing in data handling scenarios.
        /// </summary>
        /// <param name="dataType">The data type to check, represented as a string.</param>
        /// <param name="DataType">A list of strings representing the predefined basic data types.</param>
        /// <returns>
        /// True if the specified data type is recognized as a basic data type; otherwise, false.
        /// This method relies on a predefined list of data types, represented by the 'DataType' collection.
        /// </returns>
        public bool IsBasicDataType(string dataType, List<string> DataType)
        {
            return DataType.Contains(dataType);
        }

        /// <summary>
        /// Checks if the provided string matches a specific pattern. 
        /// This pattern starts with "PD:xxx:", followed by one or more digits, a sequence of letters, and concludes with a dollar sign ($).
        /// This method is useful for validating strings that are expected to adhere to a predefined format, especially in configurations or data parsing scenarios.
        /// </summary>
        /// <param name="data">The string to be checked against the regex pattern.</param>
        /// <returns>
        /// True if the string matches the pattern; otherwise, false.
        /// </returns>
        static bool CheckForPatternMatch(string data) //Check for specific Pattern for description
        {

            Regex regex = new Regex(@"^PD:xxx:(\d+)[a-zA-Z]+\$$");

            Match match = regex.Match(data);
            if (match.Success)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Replaces placeholders in a string with specified replacement values.
        /// This method is used to dynamically update specific substrings within a string, identified by placeholders "xxx" and "yy".
        /// </summary>
        /// <param name="processedString">The string in which placeholders will be replaced.</param>
        /// <param name="replacementXx">The string to replace instances of "xxx" in the processedString. If null, no replacements are made for "xxx".</param>
        /// <param name="replacementYy">The string to replace instances of "yy" in the processedString. If null, no replacements are made for "yy".</param>
        /// <returns>
        /// The modified string after the replacements have been made. If either replacement string is null, the processedString is returned without those specific replacements.
        /// </returns>
        static string ReplaceValues(string processedString, string replacementXx, string replacementYy)
        {
            processedString = replacementXx != null ? processedString.Replace("xxx", replacementXx) : string.Empty;
            processedString = replacementYy != null ? processedString.Replace("yy", replacementYy) : string.Empty;

            return processedString;
        }

        ///////////////////////////////////////JSON To File//////////////////////////////////////////////

        public void JSONToFile(OPCNode opcNode, List<OPCPage> pages)
        {
            try
            {
                var nodeInfo = GetNodeInfo(opcNode);
                var page = pages.FirstOrDefault(x => x.PageName == nodeInfo.Template);
                bool IsPageNotExist = false;
                if (nodeInfo.IsSimpleType)
                {
                    page.objects.Add(new OPCColumnType() { dataType = nodeInfo.SimpleType, objectLevel = opcNode.ObjectLevels, tagName = nodeInfo.Name, description = nodeInfo.Description });
                }
                if (nodeInfo.IsComplexType)
                {
                    foreach (var complexType in nodeInfo.ComplexTypes)
                    {
                        page?.objects.Add(new OPCColumnType() { dataType = complexType, objectLevel = opcNode.ObjectLevels, tagName = nodeInfo.Name, description = nodeInfo.Description });

                        if (!pages.Exists(x => x.PageName == complexType))
                        {
                            IsPageNotExist = true;
                            pages.Add(new OPCPage()
                            {
                                PageName = complexType
                            });
                        }
                    }
                    if (IsPageNotExist)
                    {
                        foreach (var child in opcNode.childTypes)
                        {
                            JSONToFile(child, pages);
                        }
                    }
                }
                if (nodeInfo.IsObjectLevel)
                {
                    foreach (var child in opcNode.childTypes)
                    {
                        child.ObjectLevels.AddRange(opcNode.ObjectLevels);
                        child.ObjectLevels.Add(nodeInfo.Name);
                        child.ObjectLevels.Add(nodeInfo.Description);

                        JSONToFile(child, pages);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private NodeType GetNodeInfo(OPCNode nodeData)
        {
            try
            {
                if (nodeData.childTypes == null || nodeData.childTypes.Count == 0) // Simple tag name
                {
                    return new NodeType()
                    {
                        IsSimpleType = true,
                        Name = nodeData.name,
                        Description = nodeData.description,
                        SimpleType = nodeData.type,
                        Template = nodeData.template == null ? "Empty" : nodeData.template,
                    };
                }
                //if (nodeData.childTypes.All(x => x.type == "Empty")) //Check if it is object level
                //{
                //    return new NodeType()
                //    {
                //        IsObjectLevel = true,
                //        Name = nodeData.name,
                //        Description = nodeData.description,
                //        Template = nodeData.template == null ? "Empty" : nodeData.template,
                //    };
                //}
                if (nodeData.childTypes.Any(x => x.template == nodeData.template)) //Leaf Object level
                {
                    return new NodeType()
                    {
                        IsObjectLevel = true,
                        Name = nodeData.name,
                        Description = nodeData.description,
                        Template = nodeData.template == null ? "Empty" : nodeData.template,
                    };
                }
                return new NodeType() //Complex type
                {

                    IsComplexType = true,
                    Name = nodeData.name,
                    Description = nodeData.description,
                    Template = nodeData.template == null ? "Empty" : nodeData.template,
                    ComplexTypes = nodeData.childTypes.Select(x => x.template).Distinct().ToList()
                };
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
