using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ConsoleApp1
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // Load configuration
                ConfigurationManager.LoadConfiguration();
                var configuration = ConfigurationManager.Configuration;

                string filePath = null;

                // Check if file path is provided as command line argument
                if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
                {
                    filePath = args[0];
                }
                else
                {
                    // Show file dialog for file selection
                OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Title = "Select a Calibration Report File (CSV or Excel)";
                    openFileDialog.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                    filePath = openFileDialog.FileName;
                }

                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    Console.WriteLine("Error: File not found or path is empty.");
                    return;
                }

                // Read and extract only the required data
                var extractor = new CalibrationDataExtractor();
                var extractedData = extractor.ExtractRequiredData(filePath);

                // Save data to CSV file
                var dataExporter = new DataExporter();
                dataExporter.SaveToCsv(extractedData, filePath);

                // Write to OPC UA (Kepware)
                WriteToOpcUaAsync(extractedData, configuration).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Application will close automatically after completion
            }
        }

        static async Task WriteToOpcUaAsync(ExtractedCalibrationData data, AppConfiguration configuration)
        {
            OpcUaService opcUaService = null;
            try
            {
                Console.WriteLine("🔄 Connecting to OPC UA Server (Kepware)...");
                
                opcUaService = new OpcUaService(configuration.OpcUaSettings);
                var connected = await opcUaService.ConnectAsync();

                if (!connected)
                {
                    Console.WriteLine("⚠️ Could not connect to OPC UA Server. Data will only be saved to CSV.");
                    return;
                }

                // Map data to OPC UA write items
                var nodeMappingService = new NodeMappingService(configuration.OpcUaSettings.NodeMappings);
                var writeItems = nodeMappingService.MapCalibrationDataToOpcUaItems(data);

                // Write to OPC UA
                var success = await opcUaService.WriteBatchAsync(writeItems);
                
                if (success)
                {
                    Console.WriteLine("✅ Data successfully written to OPC UA Server (Kepware)!");
                }
                else
                {
                    Console.WriteLine("⚠️ Some data may not have been written to OPC UA Server.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ OPC UA write error: {ex.Message}");
                Console.WriteLine("Data has been saved to CSV file.");
            }
            finally
            {
                if (opcUaService != null)
                {
                    await opcUaService.DisconnectAsync();
                    opcUaService.Dispose();
                }
            }
        }







    }

public class ExtractedCalibrationData
{
        public ZeroCellVolumeData ZeroCellVolume { get; set; } = new ZeroCellVolumeData();
        public VolumeCalibrationData VolumeCalibration { get; set; } = new VolumeCalibrationData();
    }

    public class ZeroCellVolumeData
    {
        public string ChamberInsert { get; set; } = "";
        public string AnalysisStart { get; set; } = "";
        public string AnalysisEnd { get; set; } = "";
        public string Temperature { get; set; } = "";
        public string NumberOfPurges { get; set; } = "";
        public string PurgeFillPressure { get; set; } = "";
        public string NumberOfCycles { get; set; } = "";
        public string CycleFillPressure { get; set; } = "";
        public string EquilibRate { get; set; } = "";
        public string ExpansionVolume { get; set; } = "";
        public List<CycleData> Cycles { get; set; } = new List<CycleData>();
        public string AverageOffset { get; set; } = "";
        public List<string> StandardDeviations { get; set; } = new List<string>();
        public string AverageCellVolume { get; set; } = "";
    }

    public class VolumeCalibrationData
    {
        public string ChamberInsert { get; set; } = "";
        public string AnalysisStart { get; set; } = "";
        public string AnalysisEnd { get; set; } = "";
        public string Temperature { get; set; } = "";
        public string Reported { get; set; } = "";
        public string VolOfCalStandard { get; set; } = "";
        public string NumberOfPurges { get; set; } = "";
        public string PurgeFillPressure { get; set; } = "";
        public string NumberOfCycles { get; set; } = "";
        public string CycleFillPressure { get; set; } = "";
        public string EquilibRate { get; set; } = "";
        public List<VolumeCalibrationCycleData> Cycles { get; set; } = new List<VolumeCalibrationCycleData>();
        public string AverageOffset { get; set; } = "";
        public List<string> StandardDeviations { get; set; } = new List<string>();
        public string AverageScaleFactor { get; set; } = "";
        public string AverageCellVolume { get; set; } = "";
        public string AverageExpansionVolume { get; set; } = "";
    }

    public class CycleData
    {
        public string CycleNumber { get; set; } = "";
        public string CellVolume { get; set; } = "";
        public string Deviation { get; set; } = "";
    }

    public class VolumeCalibrationCycleData
    {
        public string CycleNumber { get; set; } = "";
        public string CellVolume { get; set; } = "";
        public string Deviation { get; set; } = "";
        public string ExpansionVolume { get; set; } = "";
        public string ExpansionDeviation { get; set; } = "";
}

public class CalibrationDataExtractor
{
    public ExtractedCalibrationData ExtractRequiredData(string filePath)
        {
            var data = new ExtractedCalibrationData();
            
            // Check file extension to determine file type
            string extension = Path.GetExtension(filePath).ToLower();
            
            if (extension == ".xlsx" || extension == ".xls")
            {
                // Read from Excel file
                return ExtractFromExcel(filePath);
            }
            else
            {
                // Read from CSV file
                return ExtractFromCsv(filePath);
            }
        }

        private ExtractedCalibrationData ExtractFromCsv(string filePath)
        {
            var data = new ExtractedCalibrationData();
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
        
        bool inZeroCellVolumeSection = false;
        bool inVolumeCalibrationSection = false;
        bool zeroCellVolumeHeaderFound = false;
        bool volumeCalibrationHeaderFound = false;
        bool zeroCellVolumeReportFound = false;
        bool volumeCalibrationReportFound = false;

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;

            // Check if line contains pipe separator
            var pipeIndex = line.IndexOf('|');
            var leftFields = new string[0];
            var rightFields = new string[0];

            if (pipeIndex >= 0)
            {
                var leftPart = line.Substring(0, pipeIndex);
                var rightPart = line.Substring(pipeIndex + 1);
                leftFields = ParseCsvLine(leftPart);
                rightFields = ParseCsvLine(rightPart);
            }
            else
            {
                leftFields = ParseCsvLine(line);
            }

            // Check for section headers
            foreach (var field in leftFields)
            {
                if (field.IndexOf("Zero Cell Volume Header", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    zeroCellVolumeHeaderFound = true;
                    inZeroCellVolumeSection = true;
                    inVolumeCalibrationSection = false;
                }
                if (field.IndexOf("Zero Cell Volume Report", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    zeroCellVolumeReportFound = true;
                }
            }

            foreach (var field in rightFields)
            {
                if (field.IndexOf("Volume Calibration Header", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    volumeCalibrationHeaderFound = true;
                    inVolumeCalibrationSection = true;
                    inZeroCellVolumeSection = false;
                }
                if (field.IndexOf("Volume Calibration Report", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    volumeCalibrationReportFound = true;
                }
            }

            // Extract Zero Cell Volume data (left side)
            if (inZeroCellVolumeSection || zeroCellVolumeHeaderFound)
            {
                ExtractZeroCellVolumeData(leftFields, line, data.ZeroCellVolume, zeroCellVolumeReportFound);
            }

            // Extract Volume Calibration data (right side)
            if (inVolumeCalibrationSection || volumeCalibrationHeaderFound)
            {
                ExtractVolumeCalibrationData(rightFields, line, data.VolumeCalibration, volumeCalibrationReportFound);
            }
        }

        return data;
    }

        private ExtractedCalibrationData ExtractFromExcel(string filePath)
        {
        var data = new ExtractedCalibrationData();
        
        // Check if file is .xls format (old Excel format)
        string extension = Path.GetExtension(filePath).ToLower();
        if (extension == ".xls")
        {
            Console.WriteLine("⚠️ Warning: .xls files (Excel 97-2003 format) are not fully supported.");
            Console.WriteLine("   Please convert the file to .xlsx format or save it as CSV.");
            Console.WriteLine("   Attempting to read as CSV format instead...");
            Console.WriteLine();
            
            // Try to read as CSV as a fallback
            try
            {
                return ExtractFromCsv(filePath);
            }
            catch (Exception csvEx)
            {
                Console.WriteLine($"❌ Error: Could not read .xls file as CSV either: {csvEx.Message}");
                Console.WriteLine("   Solution: Please convert the .xls file to .xlsx format using Excel.");
                return data;
            }
        }
        
        try
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Get first worksheet
                
                if (worksheet?.Dimension == null)
                {
                    return data;
                }
                
                // Iterate through all rows
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    var rowData = new List<string>();
                    
                    // Get all cells in the row
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                        rowData.Add(cellValue);
                    }
                    
                    if (rowData.Count > 0)
                    {
                        // Check if row contains pipe separator
                        var rowText = string.Join(",", rowData);
                        var pipeIndex = rowText.IndexOf('|');
                        var leftFields = new string[0];
                        var rightFields = new string[0];

                        if (pipeIndex >= 0)
                        {
                            var leftPart = rowText.Substring(0, pipeIndex);
                            var rightPart = rowText.Substring(pipeIndex + 1);
                            leftFields = ParseCsvLine(leftPart);
                            rightFields = ParseCsvLine(rightPart);
                        }
                        else
                        {
                            leftFields = rowData.ToArray();
                        }

                        // Similar extraction logic as CSV
                        bool zeroCellVolumeHeaderFound = false;
                        bool volumeCalibrationHeaderFound = false;
                        bool zeroCellVolumeReportFound = false;
                        bool volumeCalibrationReportFound = false;

                        foreach (var field in leftFields)
                        {
                            if (field.IndexOf("Zero Cell Volume Header", StringComparison.OrdinalIgnoreCase) >= 0)
                                zeroCellVolumeHeaderFound = true;
                            if (field.IndexOf("Zero Cell Volume Report", StringComparison.OrdinalIgnoreCase) >= 0)
                                zeroCellVolumeReportFound = true;
                        }

                        foreach (var field in rightFields)
                        {
                            if (field.IndexOf("Volume Calibration Header", StringComparison.OrdinalIgnoreCase) >= 0)
                                volumeCalibrationHeaderFound = true;
                            if (field.IndexOf("Volume Calibration Report", StringComparison.OrdinalIgnoreCase) >= 0)
                                volumeCalibrationReportFound = true;
                        }

                        if (zeroCellVolumeHeaderFound || zeroCellVolumeReportFound)
                        {
                            ExtractZeroCellVolumeData(leftFields, rowText, data.ZeroCellVolume, zeroCellVolumeReportFound);
                        }

                        if (volumeCalibrationHeaderFound || volumeCalibrationReportFound)
                        {
                            ExtractVolumeCalibrationData(rightFields, rowText, data.VolumeCalibration, volumeCalibrationReportFound);
                        }
                    }
                }
            }
        }
        catch (IOException ex)
        {
            Console.WriteLine($"❌ IO Error reading Excel file: {ex.Message}");
            if (ex.Message.Contains("OLE compound document") || ex.Message.Contains("encrypted") || 
                ex.Message.Contains("compound document") || ex.Message.Contains("password"))
            {
                Console.WriteLine();
                Console.WriteLine("   This error typically occurs when:");
                Console.WriteLine("   1. The file is in .xls format (Excel 97-2003) - Please convert to .xlsx");
                Console.WriteLine("   2. The file is corrupted - Try opening it in Excel and saving again");
                Console.WriteLine("   3. The file is password-protected - Remove password protection");
                Console.WriteLine();
                Console.WriteLine("   Attempting to read as CSV format as fallback...");
                try
                {
                    return ExtractFromCsv(filePath);
                }
                catch (Exception csvEx)
                {
                    Console.WriteLine($"   ❌ Could not read as CSV: {csvEx.Message}");
                }
            }
        }
        catch (ArgumentException ex)
        {
            Console.WriteLine($"❌ Argument Error reading Excel file: {ex.Message}");
            if (ex.Message.Contains("OLE compound document") || ex.Message.Contains("encrypted") || 
                ex.Message.Contains("compound document") || ex.Message.Contains("password"))
            {
                Console.WriteLine();
                Console.WriteLine("   This error typically occurs when:");
                Console.WriteLine("   1. The file is in .xls format (Excel 97-2003) - Please convert to .xlsx");
                Console.WriteLine("   2. The file is corrupted - Try opening it in Excel and saving again");
                Console.WriteLine("   3. The file is password-protected - Remove password protection");
                Console.WriteLine();
                Console.WriteLine("   Attempting to read as CSV format as fallback...");
                try
                {
                    return ExtractFromCsv(filePath);
                }
                catch (Exception csvEx)
                {
                    Console.WriteLine($"   ❌ Could not read as CSV: {csvEx.Message}");
                }
            }
        }
        catch (InvalidOperationException ex)
        {
            if (ex.Message.Contains("OLE compound document") || ex.Message.Contains("encrypted") || 
                ex.Message.Contains("compound document") || ex.Message.Contains("password"))
            {
                Console.WriteLine($"❌ Error: Cannot read this Excel file format.");
                Console.WriteLine();
                Console.WriteLine("   This error typically occurs when:");
                Console.WriteLine("   1. The file is in .xls format (Excel 97-2003) - Please convert to .xlsx");
                Console.WriteLine("   2. The file is corrupted - Try opening it in Excel and saving again");
                Console.WriteLine("   3. The file is password-protected - Remove password protection");
                Console.WriteLine();
                Console.WriteLine("   Attempting to read as CSV format as fallback...");
                try
                {
                    return ExtractFromCsv(filePath);
                }
                catch (Exception csvEx)
                {
                    Console.WriteLine($"   ❌ Could not read as CSV: {csvEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"❌ Error reading Excel file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error reading Excel file: {ex.Message}");
            if (ex.Message.Contains("OLE compound document") || ex.Message.Contains("encrypted") || 
                ex.Message.Contains("compound document") || ex.Message.Contains("password"))
            {
                Console.WriteLine();
                Console.WriteLine("   This error typically occurs when:");
                Console.WriteLine("   1. The file is in .xls format (Excel 97-2003) - Please convert to .xlsx");
                Console.WriteLine("   2. The file is corrupted - Try opening it in Excel and saving again");
                Console.WriteLine("   3. The file is password-protected - Remove password protection");
                Console.WriteLine();
                Console.WriteLine("   Attempting to read as CSV format as fallback...");
                try
                {
                    return ExtractFromCsv(filePath);
                }
                catch (Exception csvEx)
                {
                    Console.WriteLine($"   ❌ Could not read as CSV: {csvEx.Message}");
                }
            }
        }
        
        return data;
    }

        private void ExtractZeroCellVolumeData(string[] fields, string line, ZeroCellVolumeData data, bool inReportSection)
        {
            if (fields == null || fields.Length == 0) return;

            // Check if this is a cycle row (first field is a number)
            if (inReportSection && fields.Length > 0)
            {
                var firstField = fields[0].Trim();
                
                // Skip header rows
                if (firstField.IndexOf("Cycle", StringComparison.OrdinalIgnoreCase) >= 0 || 
                    firstField == "Cycle#")
                {
                    return;
                }
                
                // Check if first field is a cycle number (integer)
                if (int.TryParse(firstField, out int cycleNumber))
                {
                    // This is a cycle row - parse the complete row
                    if (fields.Length >= 3)
                    {
                        var cycle = new CycleData
                        {
                            CycleNumber = firstField,
                            CellVolume = fields[1].Trim(),
                            Deviation = fields[2].Trim()
                        };
                        if (!string.IsNullOrEmpty(cycle.CellVolume))
                        {
                            data.Cycles.Add(cycle);
                        }
                    }
                    return; // Don't process cycle rows as header data
                }
            }

            // Extract Standard Deviation values first (once per line, all occurrences)
            ExtractAllStandardDeviations(fields, data.StandardDeviations);

            // Extract header information from all fields
            for (int i = 0; i < fields.Length; i++)
            {
                var field = fields[i].Trim();
                if (string.IsNullOrEmpty(field)) continue;

                // Extract header information - try with fallback to next field
                if (string.IsNullOrEmpty(data.ChamberInsert))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Chamber Insert:");
                    if (value != null) data.ChamberInsert = value;
                }
                if (string.IsNullOrEmpty(data.AnalysisStart))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Analysis Start:");
                    if (value != null) data.AnalysisStart = value;
                }
                if (string.IsNullOrEmpty(data.AnalysisEnd))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Analysis End:");
                    if (value != null) data.AnalysisEnd = value;
                }
                if (string.IsNullOrEmpty(data.Temperature))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Temperature:");
                    if (value != null) data.Temperature = EncodingHelper.FixEncoding(value);
                }
                if (string.IsNullOrEmpty(data.NumberOfPurges))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Number of Purges:");
                    if (value != null) data.NumberOfPurges = value;
                }
                if (string.IsNullOrEmpty(data.PurgeFillPressure))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Purge fill pressure:");
                    if (value != null) data.PurgeFillPressure = value;
                }
                if (string.IsNullOrEmpty(data.NumberOfCycles))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Number of cycles:");
                    if (value != null) data.NumberOfCycles = value;
                }
                if (string.IsNullOrEmpty(data.CycleFillPressure))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Cycle fill pressure:");
                    if (value != null) data.CycleFillPressure = value;
                }
                if (string.IsNullOrEmpty(data.EquilibRate))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Equilib. Rate:");
                    if (value != null) data.EquilibRate = value;
                }
                if (string.IsNullOrEmpty(data.ExpansionVolume))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Expansion Volume:");
                    if (value != null) data.ExpansionVolume = value;
                }

                // Extract summary data
                if (string.IsNullOrEmpty(data.AverageOffset))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Average Offset:");
                    if (value != null) data.AverageOffset = value;
                }
                if (string.IsNullOrEmpty(data.AverageCellVolume))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Average Cell Volume:");
                    if (value != null) data.AverageCellVolume = value;
                }
            }
        }

        private void ExtractVolumeCalibrationData(string[] fields, string line, VolumeCalibrationData data, bool inReportSection)
        {
            if (fields == null || fields.Length == 0) return;

            // Check if this is a cycle row (first field is a number)
            if (inReportSection && fields.Length > 0)
            {
                var firstField = fields[0].Trim();
                
                // Skip header rows
                if (firstField.IndexOf("Cycle", StringComparison.OrdinalIgnoreCase) >= 0 || 
                    firstField == "Cycle#")
                {
                    return;
                }
                
                // Check if first field is a cycle number (integer)
                if (int.TryParse(firstField, out int cycleNumber))
                {
                    // This is a cycle row - parse the complete row
                    if (fields.Length >= 5)
                    {
                        var cycle = new VolumeCalibrationCycleData
                        {
                            CycleNumber = firstField,
                            CellVolume = fields[1].Trim(),
                            Deviation = fields[2].Trim(),
                            ExpansionVolume = fields[3].Trim(),
                            ExpansionDeviation = fields[4].Trim()
                        };
                        if (!string.IsNullOrEmpty(cycle.CellVolume))
                        {
                            data.Cycles.Add(cycle);
                        }
                    }
                    return; // Don't process cycle rows as header data
                }
            }

            // Extract Standard Deviation values first (once per line, all occurrences)
            ExtractAllStandardDeviations(fields, data.StandardDeviations);

            // Extract header information from all fields
            for (int i = 0; i < fields.Length; i++)
            {
                var field = fields[i].Trim();
                if (string.IsNullOrEmpty(field)) continue;

                // Extract header information - try with fallback to next field
                if (string.IsNullOrEmpty(data.ChamberInsert))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Chamber Insert:");
                    if (value != null) data.ChamberInsert = value;
                }
                if (string.IsNullOrEmpty(data.AnalysisStart))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Analysis Start:");
                    if (value != null) data.AnalysisStart = value;
                }
                if (string.IsNullOrEmpty(data.AnalysisEnd))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Analysis End:");
                    if (value != null) data.AnalysisEnd = value;
                }
                if (string.IsNullOrEmpty(data.Temperature))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Temperature:");
                    if (value != null) data.Temperature = EncodingHelper.FixEncoding(value);
                }
                if (string.IsNullOrEmpty(data.Reported))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Reported:");
                    if (value != null) data.Reported = value;
                }
                if (string.IsNullOrEmpty(data.VolOfCalStandard))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Vol. of Cal. Standard:");
                    if (value != null) data.VolOfCalStandard = value;
                }
                if (string.IsNullOrEmpty(data.NumberOfPurges))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Number of Purges:");
                    if (value != null) data.NumberOfPurges = value;
                }
                if (string.IsNullOrEmpty(data.PurgeFillPressure))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Purge fill pressure:");
                    if (value != null) data.PurgeFillPressure = value;
                }
                if (string.IsNullOrEmpty(data.NumberOfCycles))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Number of cycles:");
                    if (value != null) data.NumberOfCycles = value;
                }
                if (string.IsNullOrEmpty(data.CycleFillPressure))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Cycle fill pressure:");
                    if (value != null) data.CycleFillPressure = value;
                }
                if (string.IsNullOrEmpty(data.EquilibRate))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Equilib. Rate:");
                    if (value != null) data.EquilibRate = value;
                }

                // Extract summary data
                if (string.IsNullOrEmpty(data.AverageOffset))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Average Offset:");
                    if (value != null) data.AverageOffset = value;
                }
                if (string.IsNullOrEmpty(data.AverageScaleFactor))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Average Scale Factor:");
                    if (value != null) data.AverageScaleFactor = value;
                }
                if (string.IsNullOrEmpty(data.AverageCellVolume))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Average Cell Volume:");
                    if (value != null) data.AverageCellVolume = value;
                }
                if (string.IsNullOrEmpty(data.AverageExpansionVolume))
                {
                    var value = ExtractFieldValueWithFallback(fields, i, "Average Expansion Volume:");
                    if (value != null) data.AverageExpansionVolume = value;
                }
            }
        }

        private string ExtractFieldValue(string field, string label)
        {
            var index = field.IndexOf(label, StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
            {
                var value = field.Substring(index + label.Length).Trim();
                if (!string.IsNullOrEmpty(value))
                {
                    return value;
                }
            }
            return null;
        }

        private string ExtractFieldValueWithFallback(string[] fields, int startIndex, string label)
        {
            // First, try to find the label in the current field
            for (int i = startIndex; i < fields.Length; i++)
            {
                var field = fields[i].Trim();
                if (string.IsNullOrEmpty(field)) continue;

                var index = field.IndexOf(label, StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    // Check if value is in the same field after colon
                    var value = field.Substring(index + label.Length).Trim();
                    if (!string.IsNullOrEmpty(value))
                    {
                        return value;
                    }
                    
                    // If no value in same field, check next field
                    if (i + 1 < fields.Length)
                    {
                        var nextValue = fields[i + 1].Trim();
                        if (!string.IsNullOrEmpty(nextValue))
                        {
                            return nextValue;
                        }
                    }
                }
            }
            return null;
        }

        private void ExtractAllStandardDeviations(string[] fields, List<string> standardDeviations)
        {
            // Extract all Standard Deviation values from this line
            // Prevent duplicates only within the same line, but allow same value from different lines
            // Use normalized values (trimmed) for comparison to catch duplicates with different whitespace
            var foundInThisLine = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            for (int i = 0; i < fields.Length; i++)
            {
                var field = fields[i].Trim();
                if (string.IsNullOrEmpty(field)) continue;

                var index = field.IndexOf("Standard Deviation:", StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    // Check if value is in the same field after colon
                    var value = field.Substring(index + "Standard Deviation:".Length).Trim();
                    
                    // If no value in same field, check next field
                    if (string.IsNullOrEmpty(value) && i + 1 < fields.Length)
                    {
                        value = fields[i + 1].Trim();
                    }
                    
                    // Normalize the value (trim and normalize whitespace)
                    if (!string.IsNullOrEmpty(value))
                    {
                        // Normalize whitespace - replace multiple spaces with single space
                        value = Regex.Replace(value, @"\s+", " ").Trim();
                        
                        // Only prevent duplicates within the same line
                        // If same value appears in different lines, it will be added multiple times (as per requirement)
                        if (!foundInThisLine.Contains(value))
                        {
                            standardDeviations.Add(value);
                            foundInThisLine.Add(value);
                        }
                    }
                }
            }
        }

    private bool IsTimeFormat(string field)
    {
        if (string.IsNullOrEmpty(field)) return false;

        // Check for common time formats: dd-MM-yyyy HH:mm:ss or similar
        return field.Contains("-") && field.Contains(":") && 
               (field.Contains("20") || field.Contains("19")) && // Year check
               field.Length >= 16; // Minimum length for date-time
    }

    private string GetDataWithFallback(string[] fields, params int[] columnIndices)
    {
        foreach (int index in columnIndices)
        {
            if (index >= 0 && index < fields.Length && !string.IsNullOrWhiteSpace(fields[index]))
            {
                return fields[index].Trim();
            }
        }
        return "";
    }

    private string[] ParseCsvLine(string line)
    {
        if (string.IsNullOrEmpty(line))
            return new string[0];

        List<string> result = new List<string>();
        bool inQuotes = false;
        int startIndex = 0;

        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (line[i] == ',' && !inQuotes)
            {
                string field = line.Substring(startIndex, i - startIndex);
                result.Add(CleanCsvField(field));
                startIndex = i + 1;
            }
        }

        string lastField = line.Substring(startIndex);
        result.Add(CleanCsvField(lastField));

        return result.ToArray();
    }

    private string CleanCsvField(string field)
    {
        if (string.IsNullOrEmpty(field))
            return field;

        field = field.Trim();

        if (field.Length >= 2 && field.StartsWith("\"") && field.EndsWith("\""))
        {
            field = field.Substring(1, field.Length - 2);
            field = field.Replace("\"\"", "\"");
        }

        return field;
    }
}
}
