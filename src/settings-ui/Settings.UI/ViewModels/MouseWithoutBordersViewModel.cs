// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using global::PowerToys.GPOWrapper;
using Microsoft.PowerToys.Settings.UI.Helpers;
using Microsoft.PowerToys.Settings.UI.Library;
using Microsoft.PowerToys.Settings.UI.Library.Helpers;
using Microsoft.PowerToys.Settings.UI.Library.Interfaces;

namespace Microsoft.PowerToys.Settings.UI.ViewModels
{
    public class MouseWithoutBordersViewModel : Observable
    {
        private bool _connectFieldsVisible;

        public bool ConnectFieldsVisible
        {
            get => _connectFieldsVisible;

            set
            {
                if (_connectFieldsVisible != value)
                {
                    _connectFieldsVisible = value;
                    OnPropertyChanged(nameof(ConnectFieldsVisible));
                }
            }
        }

        private string _connectSecurityKey;

        public string ConnectSecurityKey
        {
            get => _connectSecurityKey;

            set
            {
                if (_connectSecurityKey != value)
                {
                    _connectSecurityKey = value;
                    OnPropertyChanged(nameof(ConnectSecurityKey));
                }
            }
        }

        private string _connectPCName;

        public string ConnectPCName
        {
            get => _connectPCName;

            set
            {
                if (_connectPCName != value)
                {
                    _connectPCName = value;
                    OnPropertyChanged(nameof(ConnectPCName));
                }
            }
        }

        private ISettingsUtils SettingsUtils { get; set; }

        private GeneralSettings GeneralSettingsConfig { get; set; }

        private GpoRuleConfigured _enabledGpoRuleConfiguration;
        private bool _enabledStateIsGPOConfigured;
        private bool _isEnabled;

        public bool IsEnabledGpoConfigured
        {
            get => _enabledStateIsGPOConfigured;
        }

        public void SubmitNewKeyRequest()
        {
            Settings.Properties.PendingKeyGenerationRequest = default(NewKeyGenerationRequest);
            Settings.Save(SettingsUtils);
        }

        public void SubmitConnectionRequest(string pcName, string securityKey)
        {
            Settings.Properties.PendingConnectionRequest = new ConnectionRequest { PCName = pcName, SecurityKey = securityKey };
            Settings.Save(SettingsUtils);
        }

        private MouseWithoutBordersSettings Settings { get; set; }

        public MouseWithoutBordersViewModel(ISettingsUtils settingsUtils, ISettingsRepository<GeneralSettings> settingsRepository, ISettingsRepository<MouseWithoutBordersSettings> moduleSettings, Func<string, int> ipcMSGCallBackFunc)
        {
            SettingsUtils = settingsUtils;

            // To obtain the general settings configurations of PowerToys Settings.
            if (settingsRepository == null)
            {
                throw new ArgumentNullException(nameof(settingsRepository));
            }

            if (moduleSettings == null)
            {
                throw new ArgumentNullException(nameof(moduleSettings));
            }

            Settings = moduleSettings.SettingsConfig;
            GeneralSettingsConfig = settingsRepository.SettingsConfig;

            _enabledGpoRuleConfiguration = GPOWrapper.GetConfiguredAlwaysOnTopEnabledValue();
            if (_enabledGpoRuleConfiguration == GpoRuleConfigured.Disabled || _enabledGpoRuleConfiguration == GpoRuleConfigured.Enabled)
            {
                // Get the enabled state from GPO.
                _enabledStateIsGPOConfigured = true;
                _isEnabled = _enabledGpoRuleConfiguration == GpoRuleConfigured.Enabled;
            }
            else
            {
                _isEnabled = GeneralSettingsConfig.Enabled.MouseWithoutBorders;
            }

            LoadMachineMatrixString();

            // set the callback functions value to handle outgoing IPC message.
            SendConfigMSG = ipcMSGCallBackFunc;
        }

        // Loads the machine matrix, taking into account changes to the machine pool.
        private void LoadMachineMatrixString()
        {
            List<string> loadMachineMatrixString = Settings.Properties.MachineMatrixString ?? new List<string>() { string.Empty, string.Empty, string.Empty, string.Empty };

            if (loadMachineMatrixString.Count < 4)
            {
                // Current logic of MWB assumes there are always 4 slots. Any other configuration means data corruption here.
                loadMachineMatrixString = new List<string>() { string.Empty, string.Empty, string.Empty, string.Empty };
            }

            bool editedTheMatrix = false; // keep track of changes to the matrix because of changes to the available machine pool.

            if (!string.IsNullOrEmpty(Settings.Properties.MachinePool?.Value))
            {
                List<string> availableMachines = new List<string>();

                // Format of this field is "NAME1:ID1,NAME2:ID2,..."
                // Load the available machines
                foreach (string availablMachineIdPair in Settings.Properties.MachinePool.Value.Split(","))
                {
                    string availablMachineName = availablMachineIdPair.Split(':')[0];
                    availableMachines.Add(availablMachineName);
                }

                // Start by removing the machines from the matrix that are no longer available to pick.
                for (int i = 0; i < loadMachineMatrixString.Count; i++)
                {
                    if (!availableMachines.Contains(loadMachineMatrixString[i]))
                    {
                        editedTheMatrix = true;
                        loadMachineMatrixString[i] = string.Empty;
                    }
                }

                // If an available machine is not in the matrix already, fill it in the first available spot.
                foreach (string availableMachineName in availableMachines)
                {
                    if (!loadMachineMatrixString.Contains(availableMachineName))
                    {
                        int availableIndex = loadMachineMatrixString.FindIndex(name => string.IsNullOrEmpty(name));
                        if (availableIndex >= 0)
                        {
                            loadMachineMatrixString[availableIndex] = availableMachineName;
                            editedTheMatrix = true;
                        }
                    }
                }
            }

            machineMatrixString = new IndexedObservableCollection<string>(loadMachineMatrixString);

            if (editedTheMatrix)
            {
                // Set the property directly to save the new matrix right away with the new available machines.
                MachineMatrixString = machineMatrixString;
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_enabledStateIsGPOConfigured)
                {
                    // If it's GPO configured, shouldn't be able to change this state.
                    return;
                }

                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    GeneralSettingsConfig.Enabled.MouseWithoutBorders = value;
                    OnPropertyChanged(nameof(IsEnabled));

                    OutGoingGeneralSettings outgoing = new OutGoingGeneralSettings(GeneralSettingsConfig);
                    SendConfigMSG(outgoing.ToString());

                    NotifyPropertyChanged();
                }
            }
        }

        public string SecurityKey
        {
            get => Settings.Properties.SecurityKey.Value;

            set
            {
                if (value != Settings.Properties.SecurityKey.Value)
                {
                    Settings.Properties.SecurityKey.Value = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool WrapMouse
        {
            get
            {
                return Settings.Properties.WrapMouse;
            }

            set
            {
                if (Settings.Properties.WrapMouse != value)
                {
                    Settings.Properties.WrapMouse = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool ShareClipboard
        {
            get
            {
                return Settings.Properties.ShareClipboard;
            }

            set
            {
                if (Settings.Properties.ShareClipboard != value)
                {
                    Settings.Properties.ShareClipboard = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool TransferFile
        {
            get
            {
                return Settings.Properties.TransferFile;
            }

            set
            {
                if (Settings.Properties.TransferFile != value)
                {
                    Settings.Properties.TransferFile = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool HideMouseAtScreenEdge
        {
            get
            {
                return Settings.Properties.HideMouseAtScreenEdge;
            }

            set
            {
                if (Settings.Properties.HideMouseAtScreenEdge != value)
                {
                    Settings.Properties.HideMouseAtScreenEdge = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool DrawMouseCursor
        {
            get
            {
                return Settings.Properties.DrawMouseCursor;
            }

            set
            {
                if (Settings.Properties.DrawMouseCursor != value)
                {
                    Settings.Properties.DrawMouseCursor = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool ValidateRemoteMachineIP
        {
            get
            {
                return Settings.Properties.ValidateRemoteMachineIP;
            }

            set
            {
                if (Settings.Properties.ValidateRemoteMachineIP != value)
                {
                    Settings.Properties.ValidateRemoteMachineIP = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool SameSubnetOnly
        {
            get
            {
                return Settings.Properties.SameSubnetOnly;
            }

            set
            {
                if (Settings.Properties.SameSubnetOnly != value)
                {
                    Settings.Properties.SameSubnetOnly = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool BlockScreenSaverOnOtherMachines
        {
            get
            {
                return Settings.Properties.BlockScreenSaverOnOtherMachines;
            }

            set
            {
                if (Settings.Properties.BlockScreenSaverOnOtherMachines != value)
                {
                    Settings.Properties.BlockScreenSaverOnOtherMachines = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool MoveMouseRelatively
        {
            get
            {
                return Settings.Properties.MoveMouseRelatively;
            }

            set
            {
                if (Settings.Properties.MoveMouseRelatively != value)
                {
                    Settings.Properties.MoveMouseRelatively = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool BlockMouseAtScreenCorners
        {
            get
            {
                return Settings.Properties.BlockMouseAtScreenCorners;
            }

            set
            {
                if (Settings.Properties.BlockMouseAtScreenCorners != value)
                {
                    Settings.Properties.BlockMouseAtScreenCorners = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private IndexedObservableCollection<string> machineMatrixString;

        public IndexedObservableCollection<string> MachineMatrixString
        {
            get => machineMatrixString;
            set
            {
                machineMatrixString = value;
                Settings.Properties.MachineMatrixString = new List<string>(value.ToEnumerable());
                NotifyPropertyChanged();
            }
        }

        public bool ShowClipboardAndNetworkStatusMessages
        {
            get
            {
                return Settings.Properties.ShowClipboardAndNetworkStatusMessages;
            }

            set
            {
                if (Settings.Properties.ShowClipboardAndNetworkStatusMessages != value)
                {
                    Settings.Properties.ShowClipboardAndNetworkStatusMessages = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
        {
            OnPropertyChanged(propertyName);
            SettingsUtils.SaveSettings(Settings.ToJsonString(), MouseWithoutBordersSettings.ModuleName);
        }

        private Func<string, int> SendConfigMSG { get; }
    }
}
