﻿// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Linq;

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.PowerToys.Settings.UI.Library;
using Microsoft.PowerToys.Settings.UI.Library.Helpers;
using Microsoft.PowerToys.Settings.UI.Library.Interfaces;
using Microsoft.PowerToys.Settings.UI.Library.ViewModels.Commands;

namespace Microsoft.PowerToys.Settings.UI.ViewModels
{
    public class VideoConferenceViewModel : Observable
    {
        private readonly ISettingsUtils _settingsUtils;

        private VideoConferenceSettings Settings { get; set; }

        private GeneralSettings GeneralSettingsConfig { get; set; }

        private const string ModuleName = "Video Conference";

        private Func<string, int> SendConfigMSG { get; }

        private Func<Task<string>> PickFileDialog { get; }

        private string _settingsConfigFileFolder = string.Empty;

        public VideoConferenceViewModel(ISettingsUtils settingsUtils, ISettingsRepository<GeneralSettings> settingsRepository, Func<string, int> ipcMSGCallBackFunc, Func<Task<string>> pickFileDialog, string configFileSubfolder = "")
        {
            PickFileDialog = pickFileDialog;

            if (settingsRepository == null)
            {
                throw new ArgumentNullException(nameof(settingsRepository));
            }

            GeneralSettingsConfig = settingsRepository.SettingsConfig;

            _settingsUtils = settingsUtils ?? throw new ArgumentNullException(nameof(settingsUtils));

            SendConfigMSG = ipcMSGCallBackFunc;

            _settingsConfigFileFolder = configFileSubfolder;

            try
            {
                Settings = _settingsUtils.GetSettings<VideoConferenceSettings>(GetSettingsSubPath());
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch
#pragma warning restore CA1031 // Do not catch general exception types
            {
                Settings = new VideoConferenceSettings();
                _settingsUtils.SaveSettings(Settings.ToJsonString(), GetSettingsSubPath());
            }

            CameraNames = interop.CommonManaged.GetAllVideoCaptureDeviceNames();
            MicrophoneNames = interop.CommonManaged.GetAllActiveMicrophoneDeviceNames();
            MicrophoneNames.Insert(0, "[All]");

            var shouldSaveSettings = false;

            if (string.IsNullOrEmpty(Settings.Properties.SelectedCamera.Value) && CameraNames.Count != 0)
            {
                _selectedCameraIndex = 0;
                Settings.Properties.SelectedCamera.Value = CameraNames[0];
                shouldSaveSettings = true;
            }
            else
            {
                _selectedCameraIndex = CameraNames.FindIndex(name => name == Settings.Properties.SelectedCamera.Value);
            }

            if (string.IsNullOrEmpty(Settings.Properties.SelectedMicrophone.Value))
            {
                _selectedMicrophoneIndex = 0;
                Settings.Properties.SelectedMicrophone.Value = MicrophoneNames[0];
                shouldSaveSettings = true;
            }
            else
            {
                _selectedMicrophoneIndex = MicrophoneNames.FindIndex(name => name == Settings.Properties.SelectedMicrophone.Value);
            }

            _isEnabled = GeneralSettingsConfig.Enabled.VideoConference;
            _cameraAndMicrophoneMuteHotkey = Settings.Properties.MuteCameraAndMicrophoneHotkey.Value;
            _mirophoneMuteHotkey = Settings.Properties.MuteMicrophoneHotkey.Value;
            _cameraMuteHotkey = Settings.Properties.MuteCameraHotkey.Value;
            CameraImageOverlayPath = Settings.Properties.CameraOverlayImagePath.Value;
            SelectOverlayImage = new ButtonClickCommand(SelectOverlayImageAction);
            ClearOverlayImage = new ButtonClickCommand(ClearOverlayImageAction);

            _hideToolbarWhenUnmuted = Settings.Properties.HideToolbarWhenUnmuted.Value;

            switch (Settings.Properties.ToolbarPosition.Value)
            {
                case "Top left corner":
                    _toolbarPositionIndex = 0;
                    break;
                case "Top center":
                    _toolbarPositionIndex = 1;
                    break;
                case "Top right corner":
                    _toolbarPositionIndex = 2;
                    break;
                case "Bottom left corner":
                    _toolbarPositionIndex = 3;
                    break;
                case "Bottom center":
                    _toolbarPositionIndex = 4;
                    break;
                case "Bottom right corner":
                    _toolbarPositionIndex = 5;
                    break;
            }

            switch (Settings.Properties.ToolbarMonitor.Value)
            {
                case "Main monitor":
                    _toolbarMonitorIndex = 0;
                    break;

                case "All monitors":
                    _toolbarMonitorIndex = 1;
                    break;
            }

            if (shouldSaveSettings)
            {
                _settingsUtils.SaveSettings(Settings.ToJsonString(), ModuleName);
            }
        }

        private bool _isEnabled;
        private int _toolbarPositionIndex;
        private int _toolbarMonitorIndex;
        private HotkeySettings _cameraAndMicrophoneMuteHotkey;
        private HotkeySettings _mirophoneMuteHotkey;
        private HotkeySettings _cameraMuteHotkey;
        private int _selectedCameraIndex = -1;
        private int _selectedMicrophoneIndex;
        private bool _hideToolbarWhenUnmuted;

        public List<string> CameraNames { get; }

        public List<string> MicrophoneNames { get; }

        public string CameraImageOverlayPath { get; set; }

        public ButtonClickCommand SelectOverlayImage { get; set; }

        public ButtonClickCommand ClearOverlayImage { get; set; }

        private void ClearOverlayImageAction()
        {
            CameraImageOverlayPath = string.Empty;
            Settings.Properties.CameraOverlayImagePath = string.Empty;
            RaisePropertyChanged(nameof(CameraImageOverlayPath));
        }

        private async void SelectOverlayImageAction()
        {
            try
            {
                string pickedImage = await PickFileDialog().ConfigureAwait(true);
                if (pickedImage != null)
                {
                    CameraImageOverlayPath = pickedImage;
                    Settings.Properties.CameraOverlayImagePath = pickedImage;
                    RaisePropertyChanged(nameof(CameraImageOverlayPath));
                }
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch
#pragma warning restore CA1031 // Do not catch general exception types
            {
            }
        }

        public int SelectedCameraIndex
        {
            get
            {
                return _selectedCameraIndex;
            }

            set
            {
                if (_selectedCameraIndex != value)
                {
                    _selectedCameraIndex = value;
                    if (_selectedCameraIndex >= 0 && _selectedCameraIndex < CameraNames.Count)
                    {
                        Settings.Properties.SelectedCamera.Value = CameraNames[_selectedCameraIndex];
                        RaisePropertyChanged();
                    }
                }
            }
        }

        public int SelectedMicrophoneIndex
        {
            get
            {
                return _selectedMicrophoneIndex;
            }

            set
            {
                if (_selectedMicrophoneIndex != value)
                {
                    _selectedMicrophoneIndex = value;
                    if (_selectedMicrophoneIndex >= 0 && _selectedMicrophoneIndex < MicrophoneNames.Count)
                    {
                        Settings.Properties.SelectedMicrophone.Value = MicrophoneNames[_selectedMicrophoneIndex];
                        RaisePropertyChanged();
                    }
                }
            }
        }

        public bool IsEnabled
        {
            get
            {
                return _isEnabled;
            }

            set
            {
                if (value != _isEnabled)
                {
                    _isEnabled = value;
                    GeneralSettingsConfig.Enabled.VideoConference = value;
                    OutGoingGeneralSettings snd = new OutGoingGeneralSettings(GeneralSettingsConfig);

                    SendConfigMSG(snd.ToString());
                    OnPropertyChanged(nameof(IsEnabled));
                }
            }
        }

        public bool IsElevated
        {
            get
            {
                return GeneralSettingsConfig.IsElevated;
            }
        }

        public HotkeySettings CameraAndMicrophoneMuteHotkey
        {
            get
            {
                return _cameraAndMicrophoneMuteHotkey;
            }

            set
            {
                if (value != _cameraAndMicrophoneMuteHotkey)
                {
                    _cameraAndMicrophoneMuteHotkey = value;
                    Settings.Properties.MuteCameraAndMicrophoneHotkey.Value = value;
                    RaisePropertyChanged(nameof(CameraAndMicrophoneMuteHotkey));
                }
            }
        }

        public HotkeySettings MicrophoneMuteHotkey
        {
            get
            {
                return _mirophoneMuteHotkey;
            }

            set
            {
                if (value != _mirophoneMuteHotkey)
                {
                    _mirophoneMuteHotkey = value;
                    Settings.Properties.MuteMicrophoneHotkey.Value = value;
                    RaisePropertyChanged(nameof(MicrophoneMuteHotkey));
                }
            }
        }

        public HotkeySettings CameraMuteHotkey
        {
            get
            {
                return _cameraMuteHotkey;
            }

            set
            {
                if (value != _cameraMuteHotkey)
                {
                    _cameraMuteHotkey = value;
                    Settings.Properties.MuteCameraHotkey.Value = value;
                    RaisePropertyChanged(nameof(CameraMuteHotkey));
                }
            }
        }

        public int ToolbarPostionIndex
        {
            get
            {
                return _toolbarPositionIndex;
            }

            set
            {
                if (_toolbarPositionIndex != value)
                {
                    _toolbarPositionIndex = value;
                    switch (_toolbarPositionIndex)
                    {
                        case 0:
                            Settings.Properties.ToolbarPosition.Value = "Top left corner";
                            break;

                        case 1:
                            Settings.Properties.ToolbarPosition.Value = "Top center";
                            break;

                        case 2:
                            Settings.Properties.ToolbarPosition.Value = "Top right corner";
                            break;

                        case 3:
                            Settings.Properties.ToolbarPosition.Value = "Bottom left corner";
                            break;

                        case 4:
                            Settings.Properties.ToolbarPosition.Value = "Bottom center";
                            break;

                        case 5:
                            Settings.Properties.ToolbarPosition.Value = "Bottom right corner";
                            break;
                    }

                    RaisePropertyChanged(nameof(ToolbarPostionIndex));
                }
            }
        }

        public int ToolbarMonitorIndex
        {
            get
            {
                return _toolbarMonitorIndex;
            }

            set
            {
                if (_toolbarMonitorIndex != value)
                {
                    _toolbarMonitorIndex = value;
                    switch (_toolbarMonitorIndex)
                    {
                        case 0:
                            Settings.Properties.ToolbarMonitor.Value = "Main monitor";
                            break;

                        case 1:
                            Settings.Properties.ToolbarMonitor.Value = "All monitors";
                            break;
                    }

                    RaisePropertyChanged(nameof(ToolbarMonitorIndex));
                }
            }
        }

        public bool HideToolbarWhenUnmuted
        {
            get
            {
                return _hideToolbarWhenUnmuted;
            }

            set
            {
                if (value != _hideToolbarWhenUnmuted)
                {
                    _hideToolbarWhenUnmuted = value;
                    Settings.Properties.HideToolbarWhenUnmuted.Value = value;
                    RaisePropertyChanged(nameof(HideToolbarWhenUnmuted));
                }
            }
        }

        public string GetSettingsSubPath()
        {
            return _settingsConfigFileFolder + "\\" + ModuleName;
        }

#pragma warning disable CA1030 // Use events where appropriate
        public void RaisePropertyChanged([CallerMemberName] string propertyName = null)
#pragma warning restore CA1030 // Use events where appropriate
        {
            OnPropertyChanged(propertyName);
            SndVideoConferenceSettings outsettings = new SndVideoConferenceSettings(Settings);
            SndModuleSettings<SndVideoConferenceSettings> ipcMessage = new SndModuleSettings<SndVideoConferenceSettings>(outsettings);
            SendConfigMSG(ipcMessage.ToJsonString());
            _settingsUtils.SaveSettings(Settings.ToJsonString(), GetSettingsSubPath());
        }
    }

    [ComImport]
    [Guid("3E68D4BD-7135-4D10-8018-9FB6D9F33FA1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInitializeWithWindow
    {
        void Initialize(IntPtr hwnd);
    }
}
