using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Opc.Ua;
using Opc.Ua.Client;
using Opc.Ua.Configuration;

namespace ConsoleApp1
{
    public class OpcUaService : IDisposable
    {
        private Session _session;
        private ApplicationConfiguration _configuration;
        private readonly OpcUaSettings _settings;

        public bool IsConnected => _session?.Connected == true;

        public OpcUaService(OpcUaSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public async Task<bool> ConnectAsync()
        {
            try
            {
                await CreateApplicationConfiguration();
                await EstablishSession();
                
                if (IsConnected)
                {
                    Console.WriteLine("✅ Connected to OPC UA Server!");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Connection failed: {ex.Message}");
                return false;
            }
        }

        public async Task DisconnectAsync()
        {
            try
            {
                if (_session != null)
                {
                    await _session.CloseAsync();
                    _session.Dispose();
                    _session = null;
                    Console.WriteLine("🔌 Disconnected from OPC UA Server");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Disconnect error: {ex.Message}");
            }
        }

        public async Task<bool> WriteBatchAsync(IEnumerable<OpcUaWriteItem> writeItems)
        {
            if (!IsConnected)
            {
                Console.WriteLine("⚠️ Cannot write batch - not connected to OPC UA server");
                return false;
            }

            var items = writeItems?.ToList();
            if (items == null || !items.Any())
            {
                Console.WriteLine("⚠️ No items to write");
                return true;
            }

            try
            {
                Console.WriteLine($"🔄 Writing batch of {items.Count} items to OPC UA...");
                
                var writeValues = new WriteValueCollection();
                foreach (var item in items)
                {
                    writeValues.Add(CreateWriteValue(item.NodeId, item.Value));
                }

                _session.Write(null, writeValues, out StatusCodeCollection results, out DiagnosticInfoCollection diagnostics);

                var successCount = 0;
                for (int i = 0; i < results.Count && i < items.Count; i++)
                {
                    var success = StatusCode.IsGood(results[i]);
                    if (success)
                    {
                        successCount++;
                        Console.WriteLine($"✅ {items[i].Description}: Success");
                    }
                    else
                    {
                        Console.WriteLine($"❌ {items[i].Description}: Failed - {results[i]}");
                    }
                }

                var allSuccess = successCount == items.Count;
                Console.WriteLine($"📊 Batch write completed: {successCount}/{items.Count} successful");
                
                if (!allSuccess)
                {
                    LogDiagnostics(diagnostics);
                }

                return allSuccess;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Batch write failed: {ex.Message}");
                return false;
            }
        }

        private async Task CreateApplicationConfiguration()
        {
            _configuration = new ApplicationConfiguration()
            {
                ApplicationName = _settings.ApplicationName,
                ApplicationType = ApplicationType.Client,
                SecurityConfiguration = new SecurityConfiguration
                {
                    ApplicationCertificate = new CertificateIdentifier(),
                    AutoAcceptUntrustedCertificates = _settings.AutoAcceptUntrustedCertificates,
                    AddAppCertToTrustedStore = true,
                    TrustedIssuerCertificates = new CertificateTrustList
                    {
                        StoreType = "Directory",
                        StorePath = "OPC Foundation/CertificateStores/UA Certificate Authorities"
                    },
                    TrustedPeerCertificates = new CertificateTrustList
                    {
                        StoreType = "Directory",
                        StorePath = "OPC Foundation/CertificateStores/UA Applications"
                    },
                    RejectedCertificateStore = new CertificateTrustList
                    {
                        StoreType = "Directory",
                        StorePath = "OPC Foundation/CertificateStores/RejectedCertificates"
                    }
                },
                ClientConfiguration = new ClientConfiguration 
                { 
                    DefaultSessionTimeout = _settings.SessionTimeout 
                }
            };

            await _configuration.Validate(ApplicationType.Client);
        }

        private async Task EstablishSession()
        {
            var selectedEndpoint = CoreClientUtils.SelectEndpoint(_settings.EndpointUrl, useSecurity: _settings.UseSecurity);
            var endpointConfiguration = EndpointConfiguration.Create(_configuration);
            var endpoint = new ConfiguredEndpoint(null, selectedEndpoint, endpointConfiguration);

            _session = await Session.Create(
                _configuration,
                endpoint,
                false,
                $"{_settings.ApplicationName} Session",
                (uint)_settings.SessionTimeout,
                null,
                null
            );
        }

        private WriteValue CreateWriteValue(string nodeId, object value)
        {
            Variant variant;
            
            if (value is double[] doubleArray)
            {
                variant = new Variant(doubleArray);
            }
            else
            {
                var stringValue = value?.ToString() ?? string.Empty;
                variant = new Variant(stringValue);
            }

            return new WriteValue()
            {
                NodeId = new NodeId(nodeId),
                AttributeId = Attributes.Value,
                Value = new DataValue(variant)
            };
        }

        private void LogDiagnostics(DiagnosticInfoCollection diagnostics)
        {
            if (diagnostics?.Count > 0)
            {
                foreach (var diagnostic in diagnostics)
                {
                    if (diagnostic != null)
                    {
                        Console.WriteLine($"    Diagnostic: {diagnostic}");
                    }
                }
            }
        }

        public void Dispose()
        {
            try
            {
                DisconnectAsync().GetAwaiter().GetResult();
            }
            catch
            {
                // Silent cleanup
            }
        }
    }
}

