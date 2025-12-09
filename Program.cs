using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Threading.Tasks;
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

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Select a Calibration Report File (CSV or Excel)";
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string filePath = openFileDialog.FileName;

                if (!File.Exists(filePath))
                {
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
                // Silent completion
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
    public string StandardDeviation { get; set; } = "";
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
    public string StandardDeviation { get; set; } = "";
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
        var lines = File.ReadAllLines(filePath);
        
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
                if (field.Contains("Zero Cell Volume Header", StringComparison.OrdinalIgnoreCase))
                {
                    zeroCellVolumeHeaderFound = true;
                    inZeroCellVolumeSection = true;
                    inVolumeCalibrationSection = false;
                }
                if (field.Contains("Zero Cell Volume Report", StringComparison.OrdinalIgnoreCase))
                {
                    zeroCellVolumeReportFound = true;
                }
            }

            foreach (var field in rightFields)
            {
                if (field.Contains("Volume Calibration Header", StringComparison.OrdinalIgnoreCase))
                {
                    volumeCalibrationHeaderFound = true;
                    inVolumeCalibrationSection = true;
                    inZeroCellVolumeSection = false;
                }
                if (field.Contains("Volume Calibration Report", StringComparison.OrdinalIgnoreCase))
                {
                    volumeCalibrationReportFound = true;
                }
            }

            // Extract Zero Cell Volume data (left side)
            if (inZeroCellVolumeSection || zeroCellVolumeHeaderFound)
            {
                ExtractZeroCellVolumeData(leftFields, data.ZeroCellVolume, zeroCellVolumeReportFound);
            }

            // Extract Volume Calibration data (right side)
            if (inVolumeCalibrationSection || volumeCalibrationHeaderFound)
            {
                ExtractVolumeCalibrationData(rightFields, data.VolumeCalibration, volumeCalibrationReportFound);
            }
        }

        return data;
    }

    private ExtractedCalibrationData ExtractFromExcel(string filePath)
    {
        var data = new ExtractedCalibrationData();
        
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
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
                            if (field.Contains("Zero Cell Volume Header", StringComparison.OrdinalIgnoreCase))
                                zeroCellVolumeHeaderFound = true;
                            if (field.Contains("Zero Cell Volume Report", StringComparison.OrdinalIgnoreCase))
                                zeroCellVolumeReportFound = true;
                        }

                        foreach (var field in rightFields)
                        {
                            if (field.Contains("Volume Calibration Header", StringComparison.OrdinalIgnoreCase))
                                volumeCalibrationHeaderFound = true;
                            if (field.Contains("Volume Calibration Report", StringComparison.OrdinalIgnoreCase))
                                volumeCalibrationReportFound = true;
                        }

                        if (zeroCellVolumeHeaderFound || zeroCellVolumeReportFound)
                        {
                            ExtractZeroCellVolumeData(leftFields, data.ZeroCellVolume, zeroCellVolumeReportFound);
                        }

                        if (volumeCalibrationHeaderFound || volumeCalibrationReportFound)
                        {
                            ExtractVolumeCalibrationData(rightFields, data.VolumeCalibration, volumeCalibrationReportFound);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading Excel file: {ex.Message}");
        }
        
        return data;
    }

    private void ExtractZeroCellVolumeData(string[] fields, ZeroCellVolumeData data, bool inReportSection)
    {
        if (fields == null || fields.Length == 0) return;

        for (int i = 0; i < fields.Length; i++)
        {
            var field = fields[i].Trim();
            if (string.IsNullOrEmpty(field)) continue;

            // Extract header information
            ExtractFieldValue(field, "Chamber Insert:", ref data.ChamberInsert);
            ExtractFieldValue(field, "Analysis Start:", ref data.AnalysisStart);
            ExtractFieldValue(field, "Analysis End:", ref data.AnalysisEnd);
            ExtractFieldValue(field, "Temperature:", ref data.Temperature);
            ExtractFieldValue(field, "Number of Purges:", ref data.NumberOfPurges);
            ExtractFieldValue(field, "Purge fill pressure:", ref data.PurgeFillPressure);
            ExtractFieldValue(field, "Number of cycles:", ref data.NumberOfCycles);
            ExtractFieldValue(field, "Cycle fill pressure:", ref data.CycleFillPressure);
            ExtractFieldValue(field, "Equilib. Rate:", ref data.EquilibRate);
            ExtractFieldValue(field, "Expansion Volume:", ref data.ExpansionVolume);
            
            // Extract cycle data (in report section)
            if (inReportSection && i < fields.Length - 1)
            {
                if (field == "Cycle#" || field.Contains("Cycle"))
                {
                    // Skip header row
                    continue;
                }
                
                // Try to parse cycle data
                if (double.TryParse(field, out _) && i + 1 < fields.Length)
                {
                    var cycle = new CycleData
                    {
                        CycleNumber = field,
                        CellVolume = i + 1 < fields.Length ? fields[i + 1].Trim() : "",
                        Deviation = i + 2 < fields.Length ? fields[i + 2].Trim() : ""
                    };
                    if (!string.IsNullOrEmpty(cycle.CellVolume))
                    {
                        data.Cycles.Add(cycle);
                    }
                }
            }

            // Extract summary data
            ExtractFieldValue(field, "Average Offset:", ref data.AverageOffset);
            ExtractFieldValue(field, "Standard Deviation:", ref data.StandardDeviation);
            ExtractFieldValue(field, "Average Cell Volume:", ref data.AverageCellVolume);
        }
    }

    private void ExtractVolumeCalibrationData(string[] fields, VolumeCalibrationData data, bool inReportSection)
    {
        if (fields == null || fields.Length == 0) return;

        for (int i = 0; i < fields.Length; i++)
        {
            var field = fields[i].Trim();
            if (string.IsNullOrEmpty(field)) continue;

            // Extract header information
            ExtractFieldValue(field, "Chamber Insert:", ref data.ChamberInsert);
            ExtractFieldValue(field, "Analysis Start:", ref data.AnalysisStart);
            ExtractFieldValue(field, "Analysis End:", ref data.AnalysisEnd);
            ExtractFieldValue(field, "Temperature:", ref data.Temperature);
            ExtractFieldValue(field, "Reported:", ref data.Reported);
            ExtractFieldValue(field, "Vol. of Cal. Standard:", ref data.VolOfCalStandard);
            ExtractFieldValue(field, "Number of Purges:", ref data.NumberOfPurges);
            ExtractFieldValue(field, "Purge fill pressure:", ref data.PurgeFillPressure);
            ExtractFieldValue(field, "Number of cycles:", ref data.NumberOfCycles);
            ExtractFieldValue(field, "Cycle fill pressure:", ref data.CycleFillPressure);
            ExtractFieldValue(field, "Equilib. Rate:", ref data.EquilibRate);
            
            // Extract cycle data (in report section)
            if (inReportSection && i < fields.Length - 1)
            {
                if (field == "Cycle#" || field.Contains("Cycle"))
                {
                    // Skip header row
                    continue;
                }
                
                // Try to parse cycle data
                if (double.TryParse(field, out _) && i + 1 < fields.Length)
                {
                    var cycle = new VolumeCalibrationCycleData
                    {
                        CycleNumber = field,
                        CellVolume = i + 1 < fields.Length ? fields[i + 1].Trim() : "",
                        Deviation = i + 2 < fields.Length ? fields[i + 2].Trim() : "",
                        ExpansionVolume = i + 3 < fields.Length ? fields[i + 3].Trim() : "",
                        ExpansionDeviation = i + 4 < fields.Length ? fields[i + 4].Trim() : ""
                    };
                    if (!string.IsNullOrEmpty(cycle.CellVolume))
                    {
                        data.Cycles.Add(cycle);
                    }
                }
            }

            // Extract summary data
            ExtractFieldValue(field, "Average Offset:", ref data.AverageOffset);
            ExtractFieldValue(field, "Standard Deviation:", ref data.StandardDeviation);
            ExtractFieldValue(field, "Average Scale Factor:", ref data.AverageScaleFactor);
            ExtractFieldValue(field, "Average Cell Volume:", ref data.AverageCellVolume);
            ExtractFieldValue(field, "Average Expansion Volume:", ref data.AverageExpansionVolume);
        }
    }

    private void ExtractFieldValue(string field, string label, ref string target)
    {
        if (string.IsNullOrEmpty(target) && field.Contains(label, StringComparison.OrdinalIgnoreCase))
        {
            var index = field.IndexOf(label, StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
            {
                var value = field.Substring(index + label.Length).Trim();
                if (!string.IsNullOrEmpty(value))
                {
                    target = value;
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
