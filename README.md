# OPC UA Excel/CSV Reader with Kepware Integration

A .NET console application that reads calibration report data from Excel and CSV files, exports it to a structured CSV format, and writes the data to Kepware via OPC UA.

## Features

- **Multi-format Support**: Reads calibration reports from both CSV and Excel (.xlsx, .xls) files
- **Dual Header Extraction**: Extracts complete data from both:
  - Zero Cell Volume Header
  - Volume Calibration Header
- **Structured CSV Export**: Saves all extracted data to `dataExport.csv` with Zero Cell Volume data first, followed by Volume Calibration data
- **OPC UA Integration**: Writes all extracted data to Kepware server via OPC UA protocol
- **Comprehensive Configuration**: Fully configurable OPC UA node mappings via `appsettings.json`

## Requirements

- .NET Framework 4.8
- Kepware OPC UA Server (or compatible OPC UA server)
- Visual Studio 2019 or later (for development)

## NuGet Packages

- `EPPlus` (6.2.10) - For Excel file reading
- `OPCFoundation.NetStandard.Opc.Ua.Client` (1.4.372.86) - For OPC UA communication
- `Newtonsoft.Json` (13.0.3) - For configuration file parsing

## Configuration

Edit `appsettings.json` to configure:

1. **OPC UA Server Settings**:
   - `EndpointUrl`: OPC UA server endpoint (default: `opc.tcp://127.0.0.1:49320`)
   - `ApplicationName`: Client application name
   - `SessionTimeout`: Session timeout in milliseconds
   - `AutoAcceptUntrustedCertificates`: Certificate handling
   - `UseSecurity`: Enable/disable security

2. **Node Mappings**: Configure OPC UA node IDs for all data fields:
   - Zero Cell Volume Header fields (Chamber Insert, Analysis Start/End, Temperature, etc.)
   - Zero Cell Volume Cycle rows (1-10)
   - Zero Cell Volume Summary data (Average Offset, Standard Deviation, Average Cell Volume)
   - Volume Calibration Header fields
   - Volume Calibration Cycle rows (1-10)
   - Volume Calibration Summary data

## Usage

1. **Run the application**: Execute `ConsoleApp1.exe`
2. **Select file**: Choose a CSV or Excel calibration report file
3. **Automatic processing**:
   - Data is extracted from both headers
   - Exported to `dataExport.csv` in the executable directory
   - Written to Kepware via OPC UA (if connection successful)

## Data Structure

### Zero Cell Volume Header
- Chamber Insert
- Analysis Start/End
- Temperature
- Number of Purges
- Purge fill pressure
- Number of cycles
- Cycle fill pressure
- Equilib. Rate
- Expansion Volume
- Cycle data (Cycle#, Cell Volume, Deviation)
- Average Offset
- Standard Deviation
- Average Cell Volume

### Volume Calibration Header
- Chamber Insert
- Analysis Start/End
- Temperature
- Reported
- Vol. of Cal. Standard
- Number of Purges
- Purge fill pressure
- Number of cycles
- Cycle fill pressure
- Equilib. Rate
- Cycle data (Cycle#, Cell Volume, Deviation, Expansion Volume, Expansion Deviation)
- Average Offset
- Standard Deviation
- Average Scale Factor
- Average Cell Volume
- Average Expansion Volume

## File Structure

```
ConsoleApplication/
├── Program.cs                 # Main application entry point and data extraction
├── DataExporter.cs            # CSV export functionality
├── OpcUaService.cs            # OPC UA client service
├── NodeMappingService.cs      # Maps extracted data to OPC UA nodes
├── ConfigurationManager.cs    # Configuration management
├── Models.cs                  # Data models for OPC UA write items
├── appsettings.json           # OPC UA configuration and node mappings
└── test_calibration_report.csv # Sample calibration report file
```

## Output

The application generates `dataExport.csv` with the following structure:

```
=== ZERO CELL VOLUME HEADER ===
[All Zero Cell Volume data fields and cycles]

=== VOLUME CALIBRATION HEADER ===
[All Volume Calibration data fields and cycles]
```

## OPC UA Node Format

Cycle rows are written as comma-separated strings:
- Zero Cell Volume: `Cycle#,CellVolume,Deviation`
- Volume Calibration: `Cycle#,CellVolume,Deviation,ExpansionVolume,ExpansionDeviation`

## Troubleshooting

- **Connection Failed**: Verify Kepware server is running and endpoint URL is correct
- **Node Write Failed**: Check node IDs in `appsettings.json` match your Kepware tag structure
- **File Read Error**: Ensure file format matches expected calibration report structure

## License

This project is provided as-is for educational and development purposes.

## Repository

[GitHub Repository](https://github.com/agarumang/OPCUAExcelCsvReader.git)

