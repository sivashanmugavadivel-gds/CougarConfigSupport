using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CougarConfigSupport
{
    public class PLCConversion
    {
        /// <summary>
        /// Reads an Excel file and converts its content to a PLCConfiguration object.
        /// </summary>
        /// <param name="filePath">The path to the Excel file to read.</param>
        /// <returns>
        /// A PLCConfiguration object populated with the data from the Excel file.
        /// </returns>
        /// <exception cref="Exception">Thrown when a required sheet is not found or when an error occurs during reading.</exception>
        [PackageTags("filePath")]
        public PLCDictionaryDataType ExcelToJSON(string filePath)
        {
            try
            {
                CougarConfigSupport.PLCDictionaryDataType plcConfig = new CougarConfigSupport.PLCDictionaryDataType();
                FileInfo fileInfo = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    var sheetNames = new HashSet<string>();
                    foreach (var sheet in package.Workbook.Worksheets)
                    {
                        sheetNames.Add(sheet.Name);
                    }

                    if (!sheetNames.Contains("PLC_Configuration_Creator"))
                    {
                        throw new Exception("Sheet \"PLC_Configuration_Creator\" does not exist in the workbook.");
                    }

                    ExcelWorksheet configSheet = package.Workbook.Worksheets["PLC_Configuration_Creator"];
                    plcConfig = new CougarConfigSupport.PLCDictionaryDataType
                    {
                        name = configSheet.Cells[1, 2]?.Value?.ToString().Trim() ?? "null value",
                        version = configSheet.Cells[2, 2]?.Value?.ToString().Trim() ?? "null value",
                        comment = configSheet.Cells[3, 2]?.Value?.ToString().Trim() ?? "null value",
                        opcToPLCMappings = new List<CougarConfigSupport.OpcToPLCMappings>()
                    };

                    for (int row = 5; row <= configSheet.Dimension.End.Row; row++) //Initial start row is 5
                    {
                        if (configSheet.Cells[row, 1].Value == null) continue;
                        var mapping = new CougarConfigSupport.OpcToPLCMappings
                        {
                            plcType = configSheet.Cells[row, 3]?.Value?.ToString().Trim() ?? "null value",
                            plcChannel = configSheet.Cells[row + 1, 4]?.Value?.ToString().Trim() ?? "none",
                            plcModel = configSheet.Cells[row + 2, 5]?.Value?.ToString().Trim() ?? "none",
                            plcAddress = configSheet.Cells[row + 3, 3]?.Value?.ToString() ?? "Default Address",
                            refreshRate = int.TryParse(configSheet.Cells[row + 4, 3]?.Value?.ToString(), out int refreshRate) ? refreshRate : 0,
                            manualRead = bool.TryParse(configSheet.Cells[row + 5, 3]?.Value?.ToString(), out bool manualRead) && manualRead,
                            connectionTimeout = int.TryParse(configSheet.Cells[row + 6, 3]?.Value?.ToString(), out int connectionTimeout) ? connectionTimeout : 0,
                            transactionTimeout = int.TryParse(configSheet.Cells[row + 7, 3]?.Value?.ToString(), out int transactionTimeout) ? transactionTimeout : 0,
                            connectionAttempts = int.TryParse(configSheet.Cells[row + 8, 3]?.Value?.ToString(), out int connectionAttempts) ? connectionAttempts : 0,
                            enabled = bool.TryParse(configSheet.Cells[row + 9, 3]?.Value?.ToString(), out bool enabled) && enabled,
                            plcSettings1 = configSheet.Cells[row + 10, 3]?.Value?.ToString().Trim() ?? "null value",
                            plcSettings2 = configSheet.Cells[row + 11, 3]?.Value?.ToString().Trim() ?? "null value",
                            nodeMapping = new List<CougarConfigSupport.NodeMapping>()
                        };

                        string nodeMappingSheets = configSheet.Cells[row + 12, 3]?.Value?.ToString();
                        if (nodeMappingSheets != null)
                        {
                            foreach (string sheetName in nodeMappingSheets.Split(','))
                            {
                                if (!sheetNames.Contains(sheetName.Trim()))
                                {
                                    throw new Exception($"Sheet '{sheetName.Trim()}' does not exist in the workbook.");
                                }
                                ExcelWorksheet nodeMappingSheet = package.Workbook.Worksheets[sheetName.Trim()];
                                for (int nmRow = 2; nmRow <= nodeMappingSheet.Dimension.End.Row; nmRow++)
                                {
                                    if (nodeMappingSheet.Cells[nmRow, 1].Value == null) continue;

                                    var node = new CougarConfigSupport.NodeMapping
                                    {
                                        plcTag = nodeMappingSheet.Cells[nmRow, 1]?.Value?.ToString().Trim() ?? "null value",
                                        plcTagType = nodeMappingSheet.Cells[nmRow, 2]?.Value?.ToString().Trim() ?? "null value",
                                        accessType = nodeMappingSheet.Cells[nmRow, 3]?.Value?.ToString().Trim() ?? "null value",
                                        plcTagElement = int.TryParse(nodeMappingSheet.Cells[nmRow, 4]?.Value?.ToString(), out int plcTagElement) ? plcTagElement : 0,
                                        modifyValueBy10 = bool.TryParse(nodeMappingSheet.Cells[nmRow, 5]?.Value?.ToString(), out bool modifyValueBy10) && modifyValueBy10,
                                        modifyValueBy = int.TryParse(nodeMappingSheet.Cells[nmRow, 6]?.Value?.ToString(), out int modifyValueBy) ? modifyValueBy : 0,
                                        description = nodeMappingSheet.Cells[nmRow, 7]?.Value?.ToString().Trim() ?? "null value",
                                        template = sheetName.Trim(),
                                        opcNode = nodeMappingSheet.Cells[nmRow, 8]?.Value?.ToString().Trim() ?? "null value"
                                    };
                                    mapping.nodeMapping.Add(node);
                                }
                            }
                        }
                        else
                        {
                            throw new Exception($"NodeMapping is empty! in row: {row + 12} on {configSheet.Cells[row, 1].Value} with PLC_Type: {configSheet.Cells[row, 3].Value}, Please use any sheet name for nodeMapping parameter");
                        }
                        plcConfig.opcToPLCMappings.Add(mapping);

                        // Skip to the next set of mappings
                        row += 12;
                    }
                    return plcConfig;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred while reading the Excel configuration.", ex);
            }
        }

        /// <summary>
        /// Creates an Excel file from a PLCConfiguration object.
        /// </summary>
        /// <param name="plcConfig">The PLCConfiguration object to be exported to Excel.</param>
        /// <param name="filePath">The file path where the Excel file will be saved.</param>
        [PackageTags("plcConfig","filePath")]
        public void JSONToExcel(PLCDictionaryDataType plcConfig, string filePath)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    // Create the main worksheet
                    ExcelWorksheet configSheet = package.Workbook.Worksheets.Add("PLC_Configuration_Creator");

                    // Populate main worksheet
                    configSheet.Cells[1, 1].Value = "Name";
                    configSheet.Cells[1, 2].Value = GetCellValue(plcConfig.name);

                    configSheet.Cells[2, 1].Value = "Version";
                    configSheet.Cells[2, 2].Value = GetCellValue(plcConfig.version);

                    configSheet.Cells[3, 1].Value = "Comment";
                    configSheet.Cells[3, 2].Value = GetCellValue(plcConfig.comment);

                    configSheet.Cells[4, 1].Value = "opcToPLCMappings";

                    int row = 5;
                    int count = 1;
                    foreach (var mapping in plcConfig.opcToPLCMappings)
                    {
                        configSheet.Cells[row, 1].Value = $"Type_{count}";
                        configSheet.Cells[row, 2].Value = "plcType";
                        configSheet.Cells[row, 3].Value = GetCellValue(mapping.plcType);
                        configSheet.Cells[row + 1, 2].Value = "plcChannel";
                        configSheet.Cells[row + 1, 3].Value = GetCellValue(mapping.plcChannel);
                        configSheet.Cells[row + 2, 2].Value = "plcModel";
                        configSheet.Cells[row + 2, 3].Value = GetCellValue(mapping.plcModel);
                        configSheet.Cells[row + 3, 2].Value = "plcAddress";
                        configSheet.Cells[row + 3, 3].Value = GetCellValue(mapping.plcAddress);
                        configSheet.Cells[row + 4, 2].Value = "refreshRate";
                        configSheet.Cells[row + 4, 3].Value = GetCellValue(mapping.refreshRate);
                        configSheet.Cells[row + 5, 2].Value = "manualRead";
                        configSheet.Cells[row + 5, 3].Value = GetCellValue(mapping.manualRead);
                        configSheet.Cells[row + 6, 2].Value = "connectionTimeout";
                        configSheet.Cells[row + 6, 3].Value = GetCellValue(mapping.connectionTimeout);
                        configSheet.Cells[row + 7, 2].Value = "transactionTimeout";
                        configSheet.Cells[row + 7, 3].Value = GetCellValue(mapping.transactionTimeout);
                        configSheet.Cells[row + 8, 2].Value = "connectionAttempts";
                        configSheet.Cells[row + 8, 3].Value = GetCellValue(mapping.connectionAttempts);
                        configSheet.Cells[row + 9, 2].Value = "enabled";
                        configSheet.Cells[row + 9, 3].Value = GetCellValue(mapping.enabled);
                        configSheet.Cells[row + 10, 2].Value = "plcSettings1";
                        configSheet.Cells[row + 10, 3].Value = GetCellValue(mapping.plcSettings1);
                        configSheet.Cells[row + 11, 2].Value = "plcSettings2";
                        configSheet.Cells[row + 11, 3].Value = GetCellValue(mapping.plcSettings2);

                        string nodeMappingSheetNames = string.Empty;
                        if (mapping.nodeMapping.Count > 0)
                        {
                            var uniqueTemplates = new HashSet<string>();
                            foreach (var node in mapping.nodeMapping)
                            {
                                uniqueTemplates.Add(node.template ?? mapping.plcType);
                            }
                            nodeMappingSheetNames = string.Join(", ", uniqueTemplates);
                            configSheet.Cells[row + 12, 2].Value = "nodeMapping";
                            configSheet.Cells[row + 12, 3].Value = nodeMappingSheetNames;
                        }
                        foreach (var nodeSheet in nodeMappingSheetNames.Split(','))
                        {
                            ExcelWorksheet nodeMappingSheet = package.Workbook.Worksheets.Add(nodeSheet.Trim());

                            nodeMappingSheet.Cells[1, 1].Value = "plcTag";
                            nodeMappingSheet.Cells[1, 2].Value = "plcTagType";
                            nodeMappingSheet.Cells[1, 3].Value = "accessType";
                            nodeMappingSheet.Cells[1, 4].Value = "plcTagElement";
                            nodeMappingSheet.Cells[1, 5].Value = "modifyValueBy10";
                            nodeMappingSheet.Cells[1, 6].Value = "modifyValueBy";
                            nodeMappingSheet.Cells[1, 7].Value = "description";
                            nodeMappingSheet.Cells[1, 8].Value = "opcNode";

                            // Create a worksheet for each nodeMapping template if it doesn't already exist
                            foreach (var node in mapping.nodeMapping.Where(n => n.template == nodeSheet.Trim() || (n.template == null && mapping.plcType == nodeSheet.Trim())))
                            {
                                int nodeRow = nodeMappingSheet.Dimension.End.Row + 1;

                                nodeMappingSheet.Cells[nodeRow, 1].Value = GetCellValue(node.plcTag);
                                nodeMappingSheet.Cells[nodeRow, 2].Value = GetCellValue(node.plcTagType);
                                nodeMappingSheet.Cells[nodeRow, 3].Value = GetCellValue(node.accessType);
                                nodeMappingSheet.Cells[nodeRow, 4].Value = GetCellValue(node.plcTagElement);
                                nodeMappingSheet.Cells[nodeRow, 5].Value = GetCellValue(node.modifyValueBy10);
                                nodeMappingSheet.Cells[nodeRow, 6].Value = GetCellValue(node.modifyValueBy);
                                nodeMappingSheet.Cells[nodeRow, 7].Value = GetCellValue(node.description);
                                nodeMappingSheet.Cells[nodeRow, 8].Value = GetCellValue(node.opcNode);
                            }
                        }
                        row += 13; // Move to the next set of mappings
                        count += 1;
                    }
                    // Child Sheet Customization (Process each sheet only once)
                    var processedSheets = new HashSet<string>();
                    foreach (var mapping in plcConfig.opcToPLCMappings)
                    {
                        foreach (var node in mapping.nodeMapping)
                        {
                            string sheetName = node.template ?? mapping.plcType;
                            if (processedSheets.Add(sheetName))
                            {
                                ExcelWorksheet nodeMappingSheetCustom = package.Workbook.Worksheets[sheetName];
                                // AutoFit columns 1 to 8
                                for (int i = 1; i <= 8; i++)
                                {
                                    nodeMappingSheetCustom.Column(i).AutoFit();
                                }
                                nodeMappingSheetCustom.Cells[1, 1, 1, 8].AutoFilter = true;
                                nodeMappingSheetCustom.Cells[1, 1, 1, 8].Style.Font.Bold = true;
                                nodeMappingSheetCustom.Cells[1, 1, 1, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                nodeMappingSheetCustom.Cells[1, 1, 1, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#70AD47"));
                            }
                        }
                    }

                    // Main Sheet Customization
                    configSheet.Column(1).Width = 15;
                    configSheet.Column(2).Width = 22;
                    configSheet.Column(3).Width = 40;
                    configSheet.Cells[1, 2, 1, 3].Merge = true;
                    configSheet.Cells[2, 2, 2, 3].Merge = true;
                    configSheet.Cells[3, 2, 3, 3].Merge = true;
                    configSheet.Cells[4, 1, 4, 3].Merge = true;
                    configSheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    int start = 5;
                    int end = 17;
                    for (int i = 1; i < count; i++)
                    {
                        configSheet.Cells[start, 1, end, 1].Merge = true;
                        configSheet.Cells[start, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        configSheet.Cells[start, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        start += 13;
                        end += 13;
                    }

                    // Save the Excel file
                    FileInfo file = new FileInfo(filePath);
                    package.SaveAs(file);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred while creating the Excel file.", ex);
            }
        }

        /// <summary>
        /// Returns a suitable cell value based on the provided object.
        /// Converts default values (e.g., false for bool, 0 for int and double) to empty strings.
        /// </summary>
        /// <param name="value">The object value to check and convert.</param>
        /// <returns>
        /// The original value if it is not a default value; otherwise, an empty string.
        /// </returns>
        [PackageTags("value")]
        private static object GetCellValue(object value)
        {
            if (value is bool booleanValue)
            {
                return booleanValue ? (object)booleanValue : "";
            }
            if (value is int intValue)
            {
                return intValue != 0 ? (object)intValue : "";
            }
            if (value is double doubleValue)
            {
                return doubleValue != 0 ? (object)doubleValue : "";
            }
            if (value is string strValue)
            {
                return !string.IsNullOrEmpty(strValue) && !strValue.Equals("null value", StringComparison.OrdinalIgnoreCase) ? (object)strValue : "";
            }
            return value ?? "";
        }
    }
}
