using System;
using System.Collections.Generic;
using System.Linq;

namespace ConsoleApp1
{
    public class NodeMappingService
    {
        private readonly NodeMappings _nodeMappings;

        public NodeMappingService(NodeMappings nodeMappings)
        {
            _nodeMappings = nodeMappings ?? throw new ArgumentNullException(nameof(nodeMappings));
        }

        public IEnumerable<OpcUaWriteItem> MapCalibrationDataToOpcUaItems(ExtractedCalibrationData data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            var writeItems = new List<OpcUaWriteItem>();

            // Map Zero Cell Volume data
            MapZeroCellVolumeData(data.ZeroCellVolume, writeItems);

            // Map Volume Calibration data
            MapVolumeCalibrationData(data.VolumeCalibration, writeItems);

            return writeItems.Where(item => !string.IsNullOrEmpty(item.NodeId));
        }

        private void MapZeroCellVolumeData(ZeroCellVolumeData data, List<OpcUaWriteItem> writeItems)
        {
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_ChamberInsert, data.ChamberInsert, "Zero Cell Volume - Chamber Insert");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_AnalysisStart, data.AnalysisStart, "Zero Cell Volume - Analysis Start");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_AnalysisEnd, data.AnalysisEnd, "Zero Cell Volume - Analysis End");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_Temperature, data.Temperature, "Zero Cell Volume - Temperature");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_Reported, data.Reported, "Zero Cell Volume - Reported");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_NumberOfPurges, data.NumberOfPurges, "Zero Cell Volume - Number of Purges");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_PurgeFillPressure, data.PurgeFillPressure, "Zero Cell Volume - Purge Fill Pressure");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_NumberOfCycles, data.NumberOfCycles, "Zero Cell Volume - Number of Cycles");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_CycleFillPressure, data.CycleFillPressure, "Zero Cell Volume - Cycle Fill Pressure");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_EquilibRate, data.EquilibRate, "Zero Cell Volume - Equilib Rate");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_ExpansionVolume, data.ExpansionVolume, "Zero Cell Volume - Expansion Volume");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_AverageOffset, data.AverageOffset, "Zero Cell Volume - Average Offset");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_OffsetStandardDeviation, data.OffsetStandardDeviation, "Zero Cell Volume - Offset Standard Deviation");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_AverageCellVolume, data.AverageCellVolume, "Zero Cell Volume - Average Cell Volume");
            AddIfNotEmpty(writeItems, _nodeMappings.ZeroCellVolume_CellVolumeStandardDeviation, data.CellVolumeStandardDeviation, "Zero Cell Volume - Cell Volume Standard Deviation");

            // Map cycle rows
            MapCycleRows(data.Cycles, writeItems, "ZeroCellVolume", 10);
        }

        private void MapVolumeCalibrationData(VolumeCalibrationData data, List<OpcUaWriteItem> writeItems)
        {
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_ChamberInsert, data.ChamberInsert, "Volume Calibration - Chamber Insert");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_AnalysisStart, data.AnalysisStart, "Volume Calibration - Analysis Start");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_AnalysisEnd, data.AnalysisEnd, "Volume Calibration - Analysis End");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_Temperature, data.Temperature, "Volume Calibration - Temperature");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_Reported, data.Reported, "Volume Calibration - Reported");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_VolOfCalStandard, data.VolOfCalStandard, "Volume Calibration - Vol of Cal Standard");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_NumberOfPurges, data.NumberOfPurges, "Volume Calibration - Number of Purges");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_PurgeFillPressure, data.PurgeFillPressure, "Volume Calibration - Purge Fill Pressure");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_NumberOfCycles, data.NumberOfCycles, "Volume Calibration - Number of Cycles");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_CycleFillPressure, data.CycleFillPressure, "Volume Calibration - Cycle Fill Pressure");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_EquilibRate, data.EquilibRate, "Volume Calibration - Equilib Rate");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_AverageOffset, data.AverageOffset, "Volume Calibration - Average Offset");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_OffsetStandardDeviation, data.OffsetStandardDeviation, "Volume Calibration - Offset Standard Deviation");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_AverageScaleFactor, data.AverageScaleFactor, "Volume Calibration - Average Scale Factor");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_ScaleFactorStandardDeviation, data.ScaleFactorStandardDeviation, "Volume Calibration - Scale Factor Standard Deviation");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_AverageCellVolume, data.AverageCellVolume, "Volume Calibration - Average Cell Volume");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_CellVolumeStandardDeviation, data.CellVolumeStandardDeviation, "Volume Calibration - Cell Volume Standard Deviation");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_AverageExpansionVolume, data.AverageExpansionVolume, "Volume Calibration - Average Expansion Volume");
            AddIfNotEmpty(writeItems, _nodeMappings.VolumeCalibration_ExpansionVolumeStandardDeviation, data.ExpansionVolumeStandardDeviation, "Volume Calibration - Expansion Volume Standard Deviation");

            // Map cycle rows
            MapVolumeCalibrationCycleRows(data.Cycles, writeItems, "VolumeCalibration", 10);
        }

        private void MapCycleRows(List<CycleData> cycles, List<OpcUaWriteItem> writeItems, string prefix, int maxCycles)
        {
            var cycleMappings = new[]
            {
                _nodeMappings.ZeroCellVolume_CycleRow1, _nodeMappings.ZeroCellVolume_CycleRow2,
                _nodeMappings.ZeroCellVolume_CycleRow3, _nodeMappings.ZeroCellVolume_CycleRow4,
                _nodeMappings.ZeroCellVolume_CycleRow5, _nodeMappings.ZeroCellVolume_CycleRow6,
                _nodeMappings.ZeroCellVolume_CycleRow7, _nodeMappings.ZeroCellVolume_CycleRow8,
                _nodeMappings.ZeroCellVolume_CycleRow9, _nodeMappings.ZeroCellVolume_CycleRow10
            };

            var maxCount = Math.Min(cycles.Count, maxCycles);
            for (int i = 0; i < maxCount; i++)
            {
                if (i < cycleMappings.Length && !string.IsNullOrEmpty(cycleMappings[i]))
                {
                    var cycle = cycles[i];
                    var cycleString = $"{cycle.CycleNumber},{cycle.CellVolume},{cycle.Deviation}";
                    writeItems.Add(new OpcUaWriteItem(cycleMappings[i], cycleString, $"{prefix} - Cycle Row {i + 1}"));
                }
            }
        }

        private void MapVolumeCalibrationCycleRows(List<VolumeCalibrationCycleData> cycles, List<OpcUaWriteItem> writeItems, string prefix, int maxCycles)
        {
            var cycleMappings = new[]
            {
                _nodeMappings.VolumeCalibration_CycleRow1, _nodeMappings.VolumeCalibration_CycleRow2,
                _nodeMappings.VolumeCalibration_CycleRow3, _nodeMappings.VolumeCalibration_CycleRow4,
                _nodeMappings.VolumeCalibration_CycleRow5, _nodeMappings.VolumeCalibration_CycleRow6,
                _nodeMappings.VolumeCalibration_CycleRow7, _nodeMappings.VolumeCalibration_CycleRow8,
                _nodeMappings.VolumeCalibration_CycleRow9, _nodeMappings.VolumeCalibration_CycleRow10
            };

            var maxCount = Math.Min(cycles.Count, maxCycles);
            for (int i = 0; i < maxCount; i++)
            {
                if (i < cycleMappings.Length && !string.IsNullOrEmpty(cycleMappings[i]))
                {
                    var cycle = cycles[i];
                    var cycleString = $"{cycle.CycleNumber},{cycle.CellVolume},{cycle.Deviation},{cycle.ExpansionVolume},{cycle.ExpansionDeviation}";
                    writeItems.Add(new OpcUaWriteItem(cycleMappings[i], cycleString, $"{prefix} - Cycle Row {i + 1}"));
                }
            }
        }

        private void AddIfNotEmpty(List<OpcUaWriteItem> items, string nodeId, string value, string description)
        {
            if (!string.IsNullOrEmpty(nodeId) && !string.IsNullOrEmpty(value))
            {
                // Apply encoding fixes before writing to Kepware
                value = EncodingHelper.FixEncoding(value);
                items.Add(new OpcUaWriteItem(nodeId, value, description));
            }
        }
    }
}

