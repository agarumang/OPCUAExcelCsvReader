using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Reflection;

namespace ConsoleApp1
{
    public class DataExporter
    {




        public void SaveToCsv(ExtractedCalibrationData data, string originalFilePath)
        {
            try
            {
                // Get the directory where the executable is located (not the current working directory)
                string directory;
                try
                {
                    directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                }
                catch
                {
                    // Fallback to current directory if getting executable location fails
                    directory = Directory.GetCurrentDirectory();
                }
                
                string csvFilePath = Path.Combine(directory, "dataExport.csv");
                
                using (var writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
                {
                    // Write Zero Cell Volume Header section first
                    writer.WriteLine("=== ZERO CELL VOLUME HEADER ===");
                    WriteSectionData(writer, data.ZeroCellVolume);
                    
                    writer.WriteLine();
                    writer.WriteLine();
                    
                    // Write Volume Calibration Header section
                    writer.WriteLine("=== VOLUME CALIBRATION HEADER ===");
                    WriteVolumeCalibrationData(writer, data.VolumeCalibration);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving CSV: {ex.Message}");
            }
        }

        private void WriteSectionData(StreamWriter writer, ZeroCellVolumeData data)
        {
            WriteField(writer, "Chamber Insert", FixEncoding(data.ChamberInsert));
            WriteField(writer, "Analysis Start", data.AnalysisStart);
            WriteField(writer, "Analysis End", data.AnalysisEnd);
            WriteField(writer, "Temperature", FixEncoding(data.Temperature));
            WriteField(writer, "Number of Purges", data.NumberOfPurges);
            WriteField(writer, "Purge fill pressure", data.PurgeFillPressure);
            WriteField(writer, "Number of cycles", data.NumberOfCycles);
            WriteField(writer, "Cycle fill pressure", data.CycleFillPressure);
            WriteField(writer, "Equilib. Rate", data.EquilibRate);
            WriteField(writer, "Expansion Volume", FixEncoding(data.ExpansionVolume));
            
            // Write cycles
            if (data.Cycles.Count > 0)
            {
                writer.WriteLine();
                writer.WriteLine("Cycles:");
                writer.WriteLine("Cycle#,Cell Volume (cm³),Deviation (cm³)");
                foreach (var cycle in data.Cycles)
                {
                    writer.WriteLine($"{EscapeCsv(cycle.CycleNumber)},{EscapeCsv(FixEncoding(cycle.CellVolume))},{EscapeCsv(FixEncoding(cycle.Deviation))}");
                }
            }
            
            writer.WriteLine();
            WriteField(writer, "Average Offset", FixEncoding(data.AverageOffset));
            // Write all Standard Deviation values
            foreach (var stdDev in data.StandardDeviations)
            {
                WriteField(writer, "Standard Deviation", FixEncoding(stdDev));
            }
            WriteField(writer, "Average Cell Volume", FixEncoding(data.AverageCellVolume));
        }

        private void WriteVolumeCalibrationData(StreamWriter writer, VolumeCalibrationData data)
        {
            WriteField(writer, "Chamber Insert", FixEncoding(data.ChamberInsert));
            WriteField(writer, "Analysis Start", data.AnalysisStart);
            WriteField(writer, "Analysis End", data.AnalysisEnd);
            WriteField(writer, "Temperature", FixEncoding(data.Temperature));
            WriteField(writer, "Reported", data.Reported);
            WriteField(writer, "Vol. of Cal. Standard", FixEncoding(data.VolOfCalStandard));
            WriteField(writer, "Number of Purges", data.NumberOfPurges);
            WriteField(writer, "Purge fill pressure", data.PurgeFillPressure);
            WriteField(writer, "Number of cycles", data.NumberOfCycles);
            WriteField(writer, "Cycle fill pressure", data.CycleFillPressure);
            WriteField(writer, "Equilib. Rate", data.EquilibRate);
            
            // Write cycles
            if (data.Cycles.Count > 0)
            {
                writer.WriteLine();
                writer.WriteLine("Cycles:");
                writer.WriteLine("Cycle#,Cell Volume (cm³),Deviation (cm³),Expansion Volume (cm³),Deviation (cm³)");
                foreach (var cycle in data.Cycles)
                {
                    writer.WriteLine($"{EscapeCsv(cycle.CycleNumber)},{EscapeCsv(FixEncoding(cycle.CellVolume))},{EscapeCsv(FixEncoding(cycle.Deviation))},{EscapeCsv(FixEncoding(cycle.ExpansionVolume))},{EscapeCsv(FixEncoding(cycle.ExpansionDeviation))}");
                }
            }
            
            writer.WriteLine();
            WriteField(writer, "Average Offset", FixEncoding(data.AverageOffset));
            // Write all Standard Deviation values
            foreach (var stdDev in data.StandardDeviations)
            {
                WriteField(writer, "Standard Deviation", FixEncoding(stdDev));
            }
            WriteField(writer, "Average Scale Factor", data.AverageScaleFactor);
            WriteField(writer, "Average Cell Volume", FixEncoding(data.AverageCellVolume));
            WriteField(writer, "Average Expansion Volume", FixEncoding(data.AverageExpansionVolume));
        }

        private void WriteField(StreamWriter writer, string fieldName, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                writer.WriteLine($"{EscapeCsv(fieldName)},{EscapeCsv(value)}");
            }
        }

        private string FixEncoding(string value)
        {
            return EncodingHelper.FixEncoding(value);
        }

        private string EscapeCsv(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "\"\"";
            
            // Escape quotes and wrap in quotes
            value = value.Replace("\"", "\"\"");
            return $"\"{value}\"";
        }



        private string RemoveVolumeUnits(string volumeString)
        {
            try
            {
                if (string.IsNullOrEmpty(volumeString))
                    return volumeString;

                // Extract only the decimal number from the string using regex
                var match = System.Text.RegularExpressions.Regex.Match(volumeString, @"[-+]?[0-9]*\.?[0-9]+");
                
                if (match.Success)
                {
                    return match.Value;
                }
                
                // If no decimal number found, return the original string
                return volumeString;
            }
            catch
            {
                // If any error occurs, return the original string
                return volumeString;
            }
        }
    }
} 