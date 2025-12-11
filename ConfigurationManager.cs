using System;
using System.IO;
using Newtonsoft.Json;
using System.Reflection;

namespace ConsoleApp1
{
    public class AppConfiguration
    {
        public OpcUaSettings OpcUaSettings { get; set; } = new OpcUaSettings();
        public ApplicationSettings ApplicationSettings { get; set; } = new ApplicationSettings();
    }

    public class OpcUaSettings
    {
        public string EndpointUrl { get; set; } = "opc.tcp://localhost:49320";
        public string ApplicationName { get; set; } = "Calibration Data Exporter OPC UA Client";
        public int SessionTimeout { get; set; } = 60000;
        public int OperationTimeout { get; set; } = 15000;
        public bool AutoAcceptUntrustedCertificates { get; set; } = true;
        public bool UseSecurity { get; set; } = false;
        public string Username { get; set; } = "";
        public string Password { get; set; } = "";
        public string PreferredAuthenticationType { get; set; } = "Anonymous";
        public NodeMappings NodeMappings { get; set; } = new NodeMappings();
    }

    public class NodeMappings
    {
        // Zero Cell Volume mappings
        public string ZeroCellVolume_ChamberInsert { get; set; } = "";
        public string ZeroCellVolume_AnalysisStart { get; set; } = "";
        public string ZeroCellVolume_AnalysisEnd { get; set; } = "";
        public string ZeroCellVolume_Temperature { get; set; } = "";
        public string ZeroCellVolume_Reported { get; set; } = "";
        public string ZeroCellVolume_NumberOfPurges { get; set; } = "";
        public string ZeroCellVolume_PurgeFillPressure { get; set; } = "";
        public string ZeroCellVolume_NumberOfCycles { get; set; } = "";
        public string ZeroCellVolume_CycleFillPressure { get; set; } = "";
        public string ZeroCellVolume_EquilibRate { get; set; } = "";
        public string ZeroCellVolume_ExpansionVolume { get; set; } = "";
        public string ZeroCellVolume_CycleRow1 { get; set; } = "";
        public string ZeroCellVolume_CycleRow2 { get; set; } = "";
        public string ZeroCellVolume_CycleRow3 { get; set; } = "";
        public string ZeroCellVolume_CycleRow4 { get; set; } = "";
        public string ZeroCellVolume_CycleRow5 { get; set; } = "";
        public string ZeroCellVolume_CycleRow6 { get; set; } = "";
        public string ZeroCellVolume_CycleRow7 { get; set; } = "";
        public string ZeroCellVolume_CycleRow8 { get; set; } = "";
        public string ZeroCellVolume_CycleRow9 { get; set; } = "";
        public string ZeroCellVolume_CycleRow10 { get; set; } = "";
        public string ZeroCellVolume_AverageOffset { get; set; } = "";
        public string ZeroCellVolume_OffsetStandardDeviation { get; set; } = "";
        public string ZeroCellVolume_AverageCellVolume { get; set; } = "";
        public string ZeroCellVolume_CellVolumeStandardDeviation { get; set; } = "";
        
        // Volume Calibration mappings
        public string VolumeCalibration_ChamberInsert { get; set; } = "";
        public string VolumeCalibration_AnalysisStart { get; set; } = "";
        public string VolumeCalibration_AnalysisEnd { get; set; } = "";
        public string VolumeCalibration_Temperature { get; set; } = "";
        public string VolumeCalibration_Reported { get; set; } = "";
        public string VolumeCalibration_VolOfCalStandard { get; set; } = "";
        public string VolumeCalibration_NumberOfPurges { get; set; } = "";
        public string VolumeCalibration_PurgeFillPressure { get; set; } = "";
        public string VolumeCalibration_NumberOfCycles { get; set; } = "";
        public string VolumeCalibration_CycleFillPressure { get; set; } = "";
        public string VolumeCalibration_EquilibRate { get; set; } = "";
        public string VolumeCalibration_CycleRow1 { get; set; } = "";
        public string VolumeCalibration_CycleRow2 { get; set; } = "";
        public string VolumeCalibration_CycleRow3 { get; set; } = "";
        public string VolumeCalibration_CycleRow4 { get; set; } = "";
        public string VolumeCalibration_CycleRow5 { get; set; } = "";
        public string VolumeCalibration_CycleRow6 { get; set; } = "";
        public string VolumeCalibration_CycleRow7 { get; set; } = "";
        public string VolumeCalibration_CycleRow8 { get; set; } = "";
        public string VolumeCalibration_CycleRow9 { get; set; } = "";
        public string VolumeCalibration_CycleRow10 { get; set; } = "";
        public string VolumeCalibration_AverageOffset { get; set; } = "";
        public string VolumeCalibration_OffsetStandardDeviation { get; set; } = "";
        public string VolumeCalibration_AverageScaleFactor { get; set; } = "";
        public string VolumeCalibration_ScaleFactorStandardDeviation { get; set; } = "";
        public string VolumeCalibration_AverageCellVolume { get; set; } = "";
        public string VolumeCalibration_CellVolumeStandardDeviation { get; set; } = "";
        public string VolumeCalibration_AverageExpansionVolume { get; set; } = "";
        public string VolumeCalibration_ExpansionVolumeStandardDeviation { get; set; } = "";
    }

    public class ApplicationSettings
    {
        public string OutputFolderName { get; set; } = "output";
        public string CsvFileName { get; set; } = "dataExport.csv";
        public int MaxMeasurementCycles { get; set; } = 10;
    }

    public static class ConfigurationManager
    {
        private static AppConfiguration _configuration;
        private const string ConfigFileName = "appsettings.json";

        public static AppConfiguration Configuration
        {
            get
            {
                if (_configuration == null)
                {
                    LoadConfiguration();
                }
                return _configuration;
            }
        }

        public static void LoadConfiguration()
        {
            try
            {
                var configPath = GetConfigPath();
                
                if (File.Exists(configPath))
                {
                    var jsonContent = File.ReadAllText(configPath);
                    _configuration = JsonConvert.DeserializeObject<AppConfiguration>(jsonContent);
                }
                else
                {
                    _configuration = new AppConfiguration();
                    SaveConfiguration();
                }
            }
            catch (Exception)
            {
                _configuration = new AppConfiguration();
            }
        }

        public static void SaveConfiguration()
        {
            try
            {
                var configPath = GetConfigPath();
                var jsonContent = JsonConvert.SerializeObject(_configuration, Formatting.Indented);
                File.WriteAllText(configPath, jsonContent);
            }
            catch (Exception)
            {
                // Silent error handling
            }
        }

        private static string GetConfigPath()
        {
            var exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var configPath = Path.Combine(exeDirectory, ConfigFileName);
            
            if (File.Exists(configPath))
            {
                return configPath;
            }

            return Path.Combine(Directory.GetCurrentDirectory(), ConfigFileName);
        }
    }
}

