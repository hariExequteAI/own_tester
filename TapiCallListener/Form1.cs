using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Windows.Forms;
using TAPI3Lib;

namespace TapiCallListener
{
    public partial class Form1 : Form
    {
        private TAPIClass _tapi;
        private List<ITAddress> _monitoredAddresses = new List<ITAddress>();
        private ListBox _lstLog;
        private ComboBox _cmbIPs;
        private ComboBox _cmbExtensions;
        private Button _btnConnect;
        private Button _btnStart;
        private Label _lblIP;
        private TableLayoutPanel _layout;
        private bool _isTapiInitialized = false;
        private bool _isMonitoring = false;

        [DllImport("tapi32.dll", CharSet = CharSet.Auto)]
        private static extern int tapiGetLocationInfo(
            StringBuilder lpszCountryCode,
            StringBuilder lpszCityCode
        );

        public Form1()
        {
            InitializeComponent();
            SetupUI();
        }

        private void SetupUI()
        {
            this.Text = "TAPI Call Listener - Enhanced";
            this.MinimumSize = new System.Drawing.Size(800, 600);

            _layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 6,
                ColumnCount = 1,
                AutoSize = true
            };
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // IP Label
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // IP Combo
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // Connect Button
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // Extension Combo
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // Start Button
            _layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // Log ListBox

            _lblIP = new Label
            {
                Text = "Select Local IP Address:",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            };
            _layout.Controls.Add(_lblIP, 0, 0);

            _cmbIPs = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            _layout.Controls.Add(_cmbIPs, 0, 1);

            _btnConnect = new Button
            {
                Text = "Connect to TAPI",
                Dock = DockStyle.Fill,
                Height = 35,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            _btnConnect.Click += BtnConnect_Click;
            _layout.Controls.Add(_btnConnect, 0, 2);

            _cmbExtensions = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 10),
                Enabled = false
            };
            _layout.Controls.Add(_cmbExtensions, 0, 3);

            _btnStart = new Button
            {
                Text = "Start Monitoring",
                Dock = DockStyle.Fill,
                Height = 35,
                Font = new System.Drawing.Font("Segoe UI", 10),
                Enabled = false
            };
            _btnStart.Click += BtnStart_Click;
            _layout.Controls.Add(_btnStart, 0, 4);

            _lstLog = new ListBox
            {
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true,
                Font = new System.Drawing.Font("Consolas", 9)
            };
            _layout.Controls.Add(_lstLog, 0, 5);

            this.Controls.Add(_layout);

            this.Load += Form1_Load;
            this.FormClosing += Form1_FormClosing;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AddLog("🚀 TAPI Call Listener Started");

            // Populate local IPv4 addresses for selection (informational)
            _cmbIPs.Items.Clear();
            foreach (var ip in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    _cmbIPs.Items.Add(ip.ToString());
            }
            if (_cmbIPs.Items.Count > 0)
                _cmbIPs.SelectedIndex = 0;
            else
                _cmbIPs.Items.Add("No IPv4 addresses found");

            // Show country and city code (optional info)
            var countryBuf = new StringBuilder(8);
            var prefixBuf = new StringBuilder(8);
            int hr = tapiGetLocationInfo(countryBuf, prefixBuf);
            if (hr == 0)
                AddLog($"📍 Location Info - Country Code: {countryBuf}, Prefix Code: {prefixBuf}");
            else
                AddLog($"⚠️ tapiGetLocationInfo failed (HRESULT 0x{hr:X8})");
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            if (_isTapiInitialized)
            {
                AddLog("✅ Already connected to TAPI.");
                return;
            }

            string selectedIP = _cmbIPs.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedIP) || selectedIP.StartsWith("No IPv4"))
            {
                AddLog("❌ Please select a valid local IP address.");
                return;
            }

            AddLog($"🌐 Selected Local IP: {selectedIP}");
            AddLog("🔄 Initializing TAPI...");

            try
            {
                _tapi = new TAPIClass();
                _tapi.Initialize();
                AddLog("✅ TAPI initialized successfully.");

                _tapi.EventFilter = (int)(TAPI_EVENT.TE_CALLNOTIFICATION |
                          TAPI_EVENT.TE_CALLSTATE |
                          TAPI_EVENT.TE_DIGITEVENT |
                          TAPI_EVENT.TE_GENERATEEVENT |
                          TAPI_EVENT.TE_PHONEEVENT |
                          TAPI_EVENT.TE_CALLMEDIA);

                ((ITTAPIEventNotification_Event)_tapi).Event += OnTapiEvent;
                AddLog("📡 Event notifications registered.");

                // List all available TAPI addresses for diagnostics
                AddLog("📋 Available TAPI Addresses:");
                var allAddresses = new List<(string Display, string Extension, ITAddress Addr)>();
                foreach (object raw in (ITCollection)_tapi.Addresses)
                {
                    if (raw is ITAddress addr)
                    {
                        string label = $"{addr.AddressName} [{addr.DialableAddress}]";
                        allAddresses.Add((label, addr.DialableAddress, addr));
                        AddLog($"   📞 {label}");
                    }
                }

                // List available TAPI extensions (grouped by extension number)
                _cmbExtensions.Items.Clear();
                var grouped = allAddresses
                    .GroupBy(x => x.Extension)
                    .Select(g => new { Extension = g.Key, Display = g.First().Display })
                    .OrderBy(x => x.Display)
                    .ToList();

                foreach (var item in grouped)
                    _cmbExtensions.Items.Add(item.Display);

                if (_cmbExtensions.Items.Count > 0)
                {
                    _cmbExtensions.SelectedIndex = 0;
                    _cmbExtensions.Enabled = true;
                    _btnStart.Enabled = true;
                    AddLog($"✅ Found {_cmbExtensions.Items.Count} TAPI extension(s).");
                }
                else
                {
                    _cmbExtensions.Items.Add("No TAPI extensions found");
                    _cmbExtensions.Enabled = false;
                    _btnStart.Enabled = false;
                    AddLog("❌ No TAPI extensions found.");
                }

                _isTapiInitialized = true;
                _btnConnect.Text = "Connected ✅";
                _btnConnect.Enabled = false;
            }
            catch (Exception ex)
            {
                AddLog($"❌ TAPI initialization error: {ex.Message}");
                AddLog($"   Stack trace: {ex.StackTrace}");
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (!_isTapiInitialized)
            {
                AddLog("❌ Please connect to TAPI first.");
                return;
            }

            if (_isMonitoring)
            {
                AddLog("✅ Already monitoring.");
                return;
            }

            if (!IsTapiServiceRunning())
            {
                AddLog("⚠️ TAPI Service is NOT running. Please start the service.");
                return;
            }

            if (_cmbExtensions.SelectedItem == null || _cmbExtensions.SelectedItem.ToString().StartsWith("No TAPI"))
            {
                AddLog("❌ No extension selected.");
                return;
            }

            string selected = _cmbExtensions.SelectedItem.ToString();
            string extNum = selected.Substring(selected.LastIndexOf('[') + 1).TrimEnd(']');

            // Find all addresses matching the selected extension number
            _monitoredAddresses.Clear();
            foreach (object raw in (ITCollection)_tapi.Addresses)
            {
                if (raw is ITAddress addr && addr.DialableAddress == extNum)
                {
                    _monitoredAddresses.Add(addr);
                }
            }

            if (_monitoredAddresses.Count == 0)
            {
                AddLog("❌ Selected extension not found in TAPI addresses.");
                return;
            }

            try
            {
                int registeredCount = 0;
                foreach (var addr in _monitoredAddresses)
                {
                    int cookie = _tapi.RegisterCallNotifications(
                        addr,
                        true,   // Monitor incoming calls
                        true,   // Monitor outgoing calls
                        TapiConstants.TAPIMEDIATYPE_AUDIO,
                        0
                    );

                    AddLog($"🍪 Registration Cookie: {cookie}");
                    AddLog($"📡 Registered on {addr.AddressName} [{addr.DialableAddress}]");
                    AddLog($"   📋 Address Details:");
                    AddLog($"      Name: {addr.AddressName}");
                    AddLog($"      Dialable: {addr.DialableAddress}");
                    AddLog($"      Provider: {addr.ServiceProviderName}");
                    AddLog($"      State: {addr.State}");
                    registeredCount++;
                }

                _isMonitoring = true;
                AddLog($"🎯 Successfully monitoring {registeredCount} address(es) for extension {extNum}");
                AddLog("👂 Listening for incoming and outgoing calls...");
                _btnStart.Text = "Monitoring ✅";
                _btnStart.Enabled = false;
            }
            catch (Exception ex)
            {
                AddLog($"❌ Failed to register call notifications: {ex.Message}");
                AddLog($"   Stack trace: {ex.StackTrace}");
            }
        }

        private void OnTapiEvent(TAPI_EVENT tapiEvent, object pEvent)
        {
            try
            {
                AddLog($"═══ TAPI Event Received: {tapiEvent} ═══");

                switch (tapiEvent)
                {
                    case TAPI_EVENT.TE_CALLNOTIFICATION:
                        var cn = (ITCallNotificationEvent)pEvent;
                        var callN = cn.Call;

                        AddLog("📞 CALL NOTIFICATION EVENT:");
                        PrintAllCallInformation(callN, "NOTIFICATION");
                        break;

                    case TAPI_EVENT.TE_CALLSTATE:
                        var cs = (ITCallStateEvent)pEvent;
                        var callS = cs.Call;
                        CALL_STATE state = cs.State;

                        AddLog($"📱 CALL STATE EVENT: {state}");
                        PrintAllCallInformation(callS, $"STATE_{state}");
                        break;

                    case TAPI_EVENT.TE_DIGITEVENT:
                        var de = (ITDigitDetectionEvent)pEvent;
                        AddLog($"🔢 DIGIT EVENT: {de.Digit}");
                        if (de.Call != null)
                            PrintAllCallInformation(de.Call, "DIGIT");
                        break;

                    case TAPI_EVENT.TE_GENERATEEVENT:
                        var ge = (ITDigitGenerationEvent)pEvent;
                        AddLog($"🎵 GENERATE EVENT");
                        if (ge.Call != null)
                            PrintAllCallInformation(ge.Call, "GENERATE");
                        break;

                    case TAPI_EVENT.TE_PHONEEVENT:
                        AddLog($"📱 PHONE EVENT");
                        break;

                    case TAPI_EVENT.TE_CALLMEDIA:
                        var cme = (ITCallMediaEvent)pEvent;
                        AddLog($"🎧 CALL MEDIA EVENT: {cme.Event}");
                        if (cme.Call != null)
                            PrintAllCallInformation(cme.Call, "MEDIA");
                        break;

                    default:
                        AddLog($"❓ Unhandled event: {tapiEvent}");
                        break;
                }

                AddLog("═══════════════════════════════════════");
            }
            catch (Exception ex)
            {
                AddLog($"❌ Event handler error: {ex.Message}");
                AddLog($"   Stack trace: {ex.StackTrace}");
            }
        }

        private void PrintAllCallInformation(ITCallInfo call, string context)
        {
            try
            {
                AddLog($"--- Call Information ({context}) ---");

                // Try to get potential unique identifiers
                try
                {
                    // Hash code as potential unique identifier
                    int hashCode = call.GetHashCode();
                    AddLog($"🆔 Call HashCode: {hashCode}");

                    // Try to access the call handle (if available)
                    try
                    {
                        var handle = call.get_CallInfoLong(CALLINFO_LONG.CIL_CALLID);
                        AddLog($"🆔 Call ID (CIL_CALLID): {handle}");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"   Call ID not available: {ex.Message}");
                    }

                    // Try other potential ID fields
                    try
                    {
                        var completionId = call.get_CallInfoLong(CALLINFO_LONG.CIL_COMPLETIONID);
                        AddLog($"🆔 Completion ID: {completionId}");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"   Completion ID not available: {ex.Message}");
                    }

                    try
                    {
                        var relatedCallId = call.get_CallInfoLong(CALLINFO_LONG.CIL_RELATEDCALLID);
                        AddLog($"🆔 Related Call ID: {relatedCallId}");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"   Related Call ID not available: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    AddLog($"   Error getting call IDs: {ex.Message}");
                }

                // Address Information
                try
                {
                    ITAddress callAddress = call.Address;
                    if (callAddress != null)
                    {
                        AddLog($"📍 Address Name: {callAddress.AddressName}");
                        AddLog($"📍 Dialable Address: {callAddress.DialableAddress}");
                        AddLog($"📍 Service Provider: {callAddress.ServiceProviderName}");
                        AddLog($"📍 Address State: {callAddress.State}");
                    }
                }
                catch (Exception ex)
                {
                    AddLog($"   Address info error: {ex.Message}");
                }

                // All available CALLINFO_STRING values
                var stringInfos = new[]
                {
                    (CALLINFO_STRING.CIS_CALLERIDNUMBER, "Caller ID Number"),
                    (CALLINFO_STRING.CIS_CALLERIDNAME, "Caller ID Name"),
                    (CALLINFO_STRING.CIS_CALLEDIDNUMBER, "Called ID Number"),
                    (CALLINFO_STRING.CIS_CALLEDIDNAME, "Called ID Name"),
                    (CALLINFO_STRING.CIS_CONNECTEDIDNUMBER, "Connected ID Number"),
                    (CALLINFO_STRING.CIS_CONNECTEDIDNAME, "Connected ID Name"),
                    (CALLINFO_STRING.CIS_REDIRECTIONIDNUMBER, "Redirection ID Number"),
                    (CALLINFO_STRING.CIS_REDIRECTIONIDNAME, "Redirection ID Name"),
                    (CALLINFO_STRING.CIS_REDIRECTINGIDNUMBER, "Redirecting ID Number"),
                    (CALLINFO_STRING.CIS_REDIRECTINGIDNAME, "Redirecting ID Name"),
                    (CALLINFO_STRING.CIS_DISPLAYABLEADDRESS, "Displayable Address"),
                    (CALLINFO_STRING.CIS_CALLINGPARTYID, "Calling Party ID"),
                    (CALLINFO_STRING.CIS_COMMENT, "Comment"),
   
                };

                foreach (var (infoType, description) in stringInfos)
                {
                    try
                    {
                        string value = call.get_CallInfoString(infoType);
                        if (!string.IsNullOrEmpty(value))
                            AddLog($"📋 {description}: {value}");
                    }
                    catch (Exception ex)
                    {
                        // Silently skip unavailable info
                    }
                }

                // All available CALLINFO_LONG values
                var longInfos = new[]
                {
                    (CALLINFO_LONG.CIL_BEARERMODE, "Bearer Mode"),
                    (CALLINFO_LONG.CIL_CALLERIDADDRESSTYPE, "Caller ID Address Type"),
                    (CALLINFO_LONG.CIL_CALLEDIDADDRESSTYPE, "Called ID Address Type"),
                    (CALLINFO_LONG.CIL_CONNECTEDIDADDRESSTYPE, "Connected ID Address Type"),
                    (CALLINFO_LONG.CIL_REDIRECTIONIDADDRESSTYPE, "Redirection ID Address Type"),
                    (CALLINFO_LONG.CIL_REDIRECTINGIDADDRESSTYPE, "Redirecting ID Address Type"),
                    (CALLINFO_LONG.CIL_ORIGIN, "Origin"),
                    (CALLINFO_LONG.CIL_REASON, "Reason"),
                    (CALLINFO_LONG.CIL_APPSPECIFIC, "App Specific"),
                    (CALLINFO_LONG.CIL_CALLPARAMSFLAGS, "Call Params Flags"),
                    (CALLINFO_LONG.CIL_CALLTREATMENT, "Call Treatment"),
                    (CALLINFO_LONG.CIL_MINRATE, "Min Rate"),
                    (CALLINFO_LONG.CIL_MAXRATE, "Max Rate"),
                    (CALLINFO_LONG.CIL_COUNTRYCODE, "Country Code"),
                    (CALLINFO_LONG.CIL_TRUNK, "Trunk"),
                    (CALLINFO_LONG.CIL_COMPLETIONID, "Completion ID Long"),
                    (CALLINFO_LONG.CIL_NUMBEROFMONITORS, "Number of Monitors"),
                    (CALLINFO_LONG.CIL_NUMBEROFOWNERS, "Number of Owners")
                };

                foreach (var (infoType, description) in longInfos)
                {
                    try
                    {
                        int value = call.get_CallInfoLong(infoType);
                        AddLog($"🔢 {description}: {value}");
                    }
                    catch (Exception ex)
                    {
                        // Silently skip unavailable info
                    }
                }

                // Call state and privilege information
                try
                {
                    var callState = call.CallState;
                    AddLog($"📊 Call State: {callState}");
                }
                catch (Exception ex)
                {
                    AddLog($"   Call state error: {ex.Message}");
                }

                try
                {
                    var privilege = call.Privilege;
                    AddLog($"🔐 Call Privilege: {privilege}");
                }
                catch (Exception ex)
                {
                    AddLog($"   Call privilege error: {ex.Message}");
                }

              
                // Additional timestamp information
                AddLog($"⏰ Event Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");

            }
            catch (Exception ex)
            {
                AddLog($"❌ Error printing call information: {ex.Message}");
                AddLog($"   Stack trace: {ex.StackTrace}");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            AddLog("🔄 Shutting down TAPI...");
            if (_tapi != null)
            {
                try
                {
                    _tapi.Shutdown();
                    AddLog("✅ TAPI shutdown complete.");
                }
                catch (Exception ex)
                {
                    AddLog($"❌ Error during shutdown: {ex.Message}");
                }
                _tapi = null;
            }
            AddLog("👋 Application closing...");
        }

        private void AddLog(string msg)
        {
            string ts = DateTime.Now.ToString("HH:mm:ss.fff");
            string line = $"{ts} - {msg}";

            if (_lstLog.InvokeRequired)
            {
                _lstLog.Invoke(new Action(() =>
                {
                    _lstLog.Items.Add(line);
                    _lstLog.TopIndex = _lstLog.Items.Count - 1;
                }));
            }
            else
            {
                _lstLog.Items.Add(line);
                _lstLog.TopIndex = _lstLog.Items.Count - 1;
            }
        }

        private bool IsTapiServiceRunning()
        {
            try
            {
                ServiceController service = new ServiceController("TapiSrv");
                return service.Status == ServiceControllerStatus.Running;
            }
            catch
            {
                return false;
            }
        }
    }
}