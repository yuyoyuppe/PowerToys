// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;
using Microsoft.PowerToys.Settings.UI.Library;
using Microsoft.PowerToys.Settings.UI.Library.Utilities;

// <summary>
//     Application settings.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using Microsoft.Win32;

[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.Properties.Setting.Values.#LoadIntSetting(System.String,System.Int32)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.Properties.Setting.Values.#SaveSetting(System.String,System.Object)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.Properties.Setting.Values.#LoadStringSetting(System.String,System.String)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.Properties.Setting.Values.#SaveSettingQWord(System.String,System.Int64)", Justification = "Dotnet port with style preservation")]

namespace MouseWithoutBorders.Class
{
    internal class Settings
    {
        internal bool Changed;

        private readonly ISettingsUtils _settingsUtils;
        private readonly object _loadingSettingsLock = new object();
        private readonly IFileSystemWatcher _watcher;

        private MouseWithoutBordersProperties _properties;
        private MouseWithoutBordersSettings _settings;

        private void UpdateSettingsFromJson()
        {
            try
            {
                if (!_settingsUtils.SettingsExists("MouseWithoutBorders"))
                {
                    var defaultSettings = new MouseWithoutBordersSettings();
                    defaultSettings.Save(_settingsUtils);
                }

                var settings = _settingsUtils.GetSettingsOrDefault<MouseWithoutBordersSettings>("MouseWithoutBorders");
                if (settings != null)
                {
                    _settings = settings;
                    if (_properties != null)
                    {
                        // Same as in CheckBoxCircle_CheckedChanged
                        if (_properties.WrapMouse != _settings.Properties.WrapMouse)
                        {
                            Common.SendMachineMatrix();
                        }

                        // Same as CheckBoxDrawMouse_CheckedChanged
                        if (_properties.DrawMouseCursor != _settings.Properties.DrawMouseCursor && !_settings.Properties.DrawMouseCursor)
                        {
                            CustomCursor.ShowFakeMouseCursor(int.MinValue, int.MinValue);
                        }

                        if (_properties.PendingConnectionRequest != null)
                        {
                            var pcName = _properties.PendingConnectionRequest.PCName;
                            var securityKey = _properties.PendingConnectionRequest.SecurityKey;
                            _properties.PendingConnectionRequest = null;
                            SaveSettingsToJson();

                            Common.MyKey = securityKey;
                            Common.MachineMatrix = new string[Common.MAX_MACHINE] { pcName.Trim().ToUpper(CultureInfo.CurrentCulture), Common.MachineName.Trim(), string.Empty, string.Empty };

                            string[] machines = Common.MachineMatrix;
                            Common.MachinePool.Initialize(machines);

                            Common.UpdateMachinePoolStringSetting();
                        }
                    }

                    _properties = _settings.Properties;
                }
            }
            catch (IOException ex)
            {
                Logger.LogEvent($"Failed to read settings: {ex.Message}", System.Diagnostics.EventLogEntryType.Error);
            }
        }

        private void SaveSettingsToJson()
        {
            lock (_loadingSettingsLock)
            {
                try
                {
                    _settings.Save(_settingsUtils);
                }
                catch (IOException ex)
                {
                    Logger.LogEvent($"Failed to write settings: {ex.Message}", System.Diagnostics.EventLogEntryType.Error);
                }
            }
        }

        internal Settings()
        {
            _settingsUtils = new SettingsUtils();
            _watcher = Helper.GetFileWatcher("MouseWithoutBorders", "settings.json", () => UpdateSettingsFromJson());
            UpdateSettingsFromJson();
        }

        internal string Username { get; set; }

        internal bool IsMyKeyRandom { get; set; }

        internal string MachineMatrixString
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return string.Join(",", _properties.DeviceNames);
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.DeviceNames = new List<string>(value.Split(","));
                    SaveSettingsToJson();
                }
            }
        }

        internal string MachinePoolString
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.MachinePool.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    if (!value.Equals(_properties.MachinePool.Value, StringComparison.OrdinalIgnoreCase))
                    {
                        _properties.MachinePool.Value = value;
                        SaveSettingsToJson();
                    }
                }
            }
        }

        internal string MyID => Application.ProductName + " Application";

        internal string MyIDEx => Application.ProductName + " Application-Ex";

        internal bool ShareClipboard
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.ShareClipboard;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.ShareClipboard = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool TransferFile
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.TransferFile;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.TransferFile = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool MatrixOneRow
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.MatrixOneRow;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.MatrixOneRow = true;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool MatrixCircle
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.WrapMouse;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.WrapMouse = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int EasyMouse
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.EasyMouse.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    SaveSettingsToJson();
                    _properties.EasyMouse.Value = value;
                }
            }
        }

        internal bool BlockMouseAtCorners
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.BlockMouseAtScreenCorners;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.BlockMouseAtScreenCorners = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal string Enc(string st, bool dec, DataProtectionScope psc)
        {
            if (st == null || st.Length < 1)
            {
                return string.Empty;
            }

            byte[] ep = Common.GetBytesU(st);
            byte[] rv, sts;

            if (dec)
            {
                sts = Convert.FromBase64String(st);
                rv = ProtectedData.Unprotect(sts, ep, psc);
                return Common.GetStringU(rv);
            }
            else
            {
                sts = Common.GetBytesU(st);
                rv = ProtectedData.Protect(sts, ep, psc);
                return Convert.ToBase64String(rv);
            }
        }

        internal string MyKey
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    if (_properties.SecurityKey.Value.Length != 0)
                    {
                        Common.Log("GETSECKEY: Key was already loaded/set: " + _properties.SecurityKey.Value);
                        return _properties.SecurityKey.Value;
                    }
                    else
                    {
                        string randomKey = Common.CreateDefaultKey();
                        _properties.SecurityKey.Value = randomKey;

                        return randomKey;
                    }
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.SecurityKey.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int MyKeyDaysToExpire
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return int.MaxValue; // TODO(@yuyoyuppe): do we still need expiration mechanics now?
                }
            }
        }

        internal bool DisableCAD
        {
            get
            {
                return false;
            }
        }

        internal bool HideLogonLogo
        {
            get
            {
                return false;
            }
        }

        internal bool HideMouse
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.HideMouseAtScreenEdge;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.HideMouseAtScreenEdge = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool BlockScreenSaver
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.BlockScreenSaverOnOtherMachines;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.BlockScreenSaverOnOtherMachines = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool BlockScreenSaverEx
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.BlockScreenSaverOnOtherMachines;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.BlockScreenSaverOnOtherMachines = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool MoveMouseRelatively
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.MoveMouseRelatively;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.MoveMouseRelatively = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal string LastPersonalizeLogonScr
        {
            get
            {
                return string.Empty;
            }
        }

        internal uint DesMachineID
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return (uint)_properties.MachineID.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.MachineID.Value = (int)value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int LastX
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.LastX.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    Common.LastX = value;
                    _properties.LastX.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int LastY
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.LastY.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    Common.LastY = value;
                    _properties.LastY.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int PackageID
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.PackageID.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.PackageID.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool FirstRun
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.FirstRun;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.FirstRun = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int HotKeySwitchMachine
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.HotkeySwitchMachine.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.HotkeySwitchMachine.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int HotKeyToggleEasyMouse
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.EasyMouse.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.EasyMouse.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal int HotKeyLockMachine
        {
            get
            {
                return 'L';
            }
        }

        internal int HotKeyReconnect
        {
            get
            {
                return 'R';
            }
        }

        internal int HotKeyCaptureScreen
        {
            get
            {
                return 'S';
            }
        }

        internal int HotKeyExitMM
        {
            get
            {
                return 'Q';
            }
        }

        internal int HotKeySwitch2AllPC
        {
            get
            {
                return 0;
            }
        }

        private int switchCount = 0;

        internal int SwitchCount
        {
            get
            {
                return switchCount;
            }

            set
            {
                switchCount = value;
                SaveSettingsToJson();
            }
        }

        internal int DumpObjectsLevel => 6;

        internal int TcpPort => _properties.TCPPort.Value;

        internal bool DrawMouse
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.DrawMouseCursor;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.DrawMouseCursor = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool DrawMouseEx
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.DrawMouseEx;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.DrawMouseEx = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool ReverseLookup
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.ValidateRemoteMachineIP;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.ValidateRemoteMachineIP = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool SameSubNetOnly
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.SameSubnetOnly;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.SameSubnetOnly = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal string Name2IP
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.Name2IP.Value;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.Name2IP.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool UseVKMap
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.UseVKMap;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.UseVKMap = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool FisrtCtrlShiftS
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.FisrtCtrlShiftS;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.FisrtCtrlShiftS = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal Hashtable VKMap
        {
            get
            {
                return new Hashtable();
            }
        }

        internal bool StealFocusWhenSwitchingMachine => _properties.StealFocusWhenSwitchingMachine;

        private string deviceId;

        internal string DeviceId
        {
            get
            {
                string newGuid = Guid.NewGuid().ToString();

                if (deviceId == null || deviceId.Length != newGuid.Length)
                {
                    string defaultId = newGuid;
                    lock (_loadingSettingsLock)
                    {
                        _properties.DeviceID = defaultId;
                        deviceId = _properties.DeviceID.Value;

                        if (deviceId.Equals(defaultId, StringComparison.OrdinalIgnoreCase))
                        {
                            return _properties.DeviceID.Value;
                        }
                    }
                }

                return deviceId;
            }
        }

        private int? machineId;

        internal int MachineId
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    machineId ??= (machineId = _properties.MachineID.Value).Value;

                    if (machineId == 0)
                    {
                        _properties.MachineID.Value = Common.Ran.Next();
                        machineId = _properties.MachineID.Value;
                    }
                }

                return machineId.Value;
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.MachineID.Value = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool OneWayControlMode => false;

        internal bool OneWayClipboardMode => false;

        internal bool ShowClipNetStatus
        {
            get
            {
                lock (_loadingSettingsLock)
                {
                    return _properties.ShowClipboardAndNetworkStatusMessages;
                }
            }

            set
            {
                lock (_loadingSettingsLock)
                {
                    _properties.ShowClipboardAndNetworkStatusMessages = value;
                    SaveSettingsToJson();
                }
            }
        }

        internal bool SendErrorLogV2
        {
            get
            {
                return false;
            }
        }
    }

    public static class Setting
    {
        internal static Settings Values = new Settings();
    }
}
