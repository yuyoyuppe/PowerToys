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
                    _properties = settings.Properties;
                }
            }
            catch (IOException ex)
            {
                Logger.LogEvent($"Failed to read settings: {ex.Message}", System.Diagnostics.EventLogEntryType.Error);
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
            get => string.Join(",", _properties.DeviceNames);
            set => _properties.DeviceNames = new List<string>(value.Split(","));
        }

        private string machinePoolString = string.Empty;

        internal string MachinePoolString
        {
            get => _properties.MachinePool.Value;

            set
            {
                if (!value.Equals(machinePoolString, StringComparison.OrdinalIgnoreCase))
                {
                    _properties.MachinePool = value;
                }
            }
        }

        internal string MyID => Application.ProductName + " Application";

        internal string MyIDEx => Application.ProductName + " Application-Ex";

        internal bool ShareClipboard
        {
            get => _properties.ShareClipboard;
            set => _properties.ShareClipboard = value;
        }

        internal bool TransferFile
        {
            get => _properties.TransferFile;

            set => _properties.TransferFile = value;
        }

        internal bool MatrixOneRow
        {
            get => _properties.MatrixOneRow;

            set => _properties.MatrixOneRow = true;
        }

        internal bool MatrixCircle
        {
            get => _properties.MatrixCircle;

            set => _properties.MatrixCircle = value;
        }

        internal int EasyMouse
        {
            get => _properties.EasyMouse.Value;

            set => _properties.EasyMouse = value;
        }

        internal bool BlockMouseAtConrners
        {
            get => _properties.BlockMouseAtScreenCorners;

            set => _properties.BlockMouseAtScreenCorners = value;
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
                if (_properties.SecurityKey.Value.Length == 0)
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

            set
            {
                _properties.SecurityKey = value;
            }
        }

        internal int MyKeyDaysToExpire
        {
            get
            {
                return int.MinValue;
            }
        }

        internal bool DisableCAD
        {
            get => false;
        }

        internal bool HideLogonLogo
        {
            get => false;
        }

        internal bool HideMouse
        {
            get => _properties.HideMouseAtScreenEdge;

            set => _properties.HideMouseAtScreenEdge = value;
        }

        internal bool BlockScreenSaver
        {
            get => _properties.BlockScreenSaverOnOtherMachines;

            set => _properties.BlockScreenSaverOnOtherMachines = value;
        }

        internal bool BlockScreenSaverEx
        {
            get => _properties.BlockScreenSaverOnOtherMachines;

            set => _properties.BlockScreenSaverOnOtherMachines = value;
        }

        internal bool MoveMouseRelatively
        {
            get => _properties.MoveMouseRelatively;

            set => _properties.MoveMouseRelatively = value;
        }

        internal string LastPersonalizeLogonScr
        {
            get => string.Empty;
        }

        internal uint DesMachineID
        {
            get => (uint)_properties.MachineID.Value;

            set => _properties.MachineID = value;
        }

        internal int LastX
        {
            get => _properties.LastX.Value;

            set
            {
                Common.LastX = value;
                _properties.LastX = value;
            }
        }

        internal int LastY
        {
            get => _properties.LastY.Value;

            set
            {
                Common.LastY = value;
                _properties.LastY = value;
            }
        }

        internal int PackageID
        {
            get => _properties.PackageID.Value;

            set => _properties.PackageID.Value = value;
        }

        internal bool FirstRun
        {
            get => _properties.FirstRun;

            set
            {
                _properties.FirstRun = value;
            }
        }

        internal int HotKeySwitchMachine
        {
            get => _properties.HotkeySwitchMachine.Value;

            set
            {
                _properties.HotkeySwitchMachine = value;
            }
        }

        internal int HotKeyToggleEasyMouse
        {
            get => _properties.EasyMouse.Value;

            set
            {
                _properties.EasyMouse.Value = value;
            }
        }

        internal int HotKeyLockMachine
        {
            get => 'L';
        }

        internal int HotKeyReconnect
        {
            get => 'R';
        }

        internal int HotKeyCaptureScreen
        {
            get => 'S';
        }

        internal int HotKeyExitMM
        {
            get => 'Q';
        }

        internal int HotKeySwitch2AllPC
        {
            get => 0;
        }

        private int switchCount = 0;

        internal int SwitchCount
        {
            get => switchCount;

            set => switchCount = value;
        }

        internal int DumpObjectsLevel => 6;

        internal int TcpPort => _properties.TCPPort.Value;

        internal bool DrawMouse
        {
            get => _properties.DrawMouseCursor;

            set => _properties.DrawMouseCursor = value;
        }

        internal bool DrawMouseEx
        {
            get => _properties.DrawMouseEx;

            set => _properties.DrawMouseEx = value;
        }

        internal bool ReverseLookup
        {
            get => _properties.ReverseLookup;

            set => _properties.ReverseLookup = value;
        }

        internal bool SameSubNetOnly
        {
            get => _properties.SameSubnetOnly;

            set => _properties.SameSubnetOnly = value;
        }

        internal string Name2IP
        {
            get => _properties.Name2IP.Value;

            set
            {
                _properties.Name2IP.Value = value;
            }
        }

        internal bool UseVKMap
        {
            get => _properties.UseVKMap;

            set
            {
                _properties.UseVKMap = value;
            }
        }

        internal bool FisrtCtrlShiftS
        {
            get => _properties.FisrtCtrlShiftS;

            set => _properties.FisrtCtrlShiftS = value;
        }

        internal Hashtable VKMap
        {
            get => new Hashtable();
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
                    _properties.DeviceID = defaultId;
                    deviceId = _properties.DeviceID.Value;

                    if (deviceId.Equals(defaultId, StringComparison.OrdinalIgnoreCase))
                    {
                        return _properties.DeviceID.Value;
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
                machineId ??= (machineId = _properties.MachineID.Value).Value;

                if (machineId == 0)
                {
                    _properties.MachineID.Value = Common.Ran.Next();
                    machineId = _properties.MachineID.Value;
                }

                return machineId.Value;
            }

            set => _properties.MachineID.Value = value;
        }

        internal bool OneWayControlMode => false;

        internal bool OneWayClipboardMode => false;

        internal bool ShowClipNetStatus
        {
            get => _properties.ShowClipNetStatus;

            set => _properties.ShowClipNetStatus = value;
        }

        internal bool SendErrorLogV2
        {
            get => false;
        }

        internal void ForceUpdateValuesFromRegistry()
        {
        }

        private void RegenerateKey(Exception e)
        {
            Common.Log(e);
            Common.KeyCorrupted = true;
            MyKey = Common.MyKey = Common.CreateRandomKey();
        }
    }

    public static class Setting
    {
        internal static Settings Values = new Settings();
    }
}
