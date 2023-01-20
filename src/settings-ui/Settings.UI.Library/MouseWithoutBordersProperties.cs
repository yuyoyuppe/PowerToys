// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace Microsoft.PowerToys.Settings.UI.Library
{
    public struct ConnectionRequest
    {
        public string PCName;
        public string SecurityKey;
    }

    public struct NewKeyGenerationRequest
    {
    }

    public class MouseWithoutBordersProperties
    {
        public StringProperty SecurityKey { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool WrapMouse { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool ShareClipboard { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool TransferFile { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool HideMouseAtScreenEdge { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool DrawMouseCursor { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool ValidateRemoteMachineIP { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool SameSubnetOnly { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool BlockScreenSaverOnOtherMachines { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool MoveMouseRelatively { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool BlockMouseAtScreenCorners { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool ShowClipboardAndNetworkStatusMessages { get; set; }

        public List<string> DeviceNames { get; set; }

        public StringProperty MachinePool { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool MatrixOneRow { get; set; }

        public IntProperty EasyMouse { get; set; }

        public IntProperty MachineID { get; set; }

        public IntProperty LastX { get; set; }

        public IntProperty LastY { get; set; }

        public IntProperty PackageID { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool FirstRun { get; set; }

        public IntProperty HotkeySwitchMachine { get; set; }

        public IntProperty TCPPort { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool DrawMouseEx { get; set; }

        public StringProperty Name2IP { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool UseVKMap { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool FisrtCtrlShiftS { get; set; }

        [JsonConverter(typeof(BoolPropertyJsonConverter))]
        public bool StealFocusWhenSwitchingMachine { get; set; }

        public StringProperty DeviceID { get; set; }

        public ConnectionRequest? PendingConnectionRequest { get; set; }

        public NewKeyGenerationRequest? PendingKeyGenerationRequest { get; set; }

        public MouseWithoutBordersProperties()
        {
            SecurityKey = new StringProperty(string.Empty);
            WrapMouse = false;
            ShareClipboard = true;
            TransferFile = true;
            HideMouseAtScreenEdge = true;
            DrawMouseCursor = true;
            ValidateRemoteMachineIP = false;
            SameSubnetOnly = false;
            BlockScreenSaverOnOtherMachines = true;
            MoveMouseRelatively = false;
            BlockMouseAtScreenCorners = false;
            ShowClipboardAndNetworkStatusMessages = false;
            EasyMouse = new IntProperty(1);
            DeviceNames = new List<string>();
            DeviceID = new StringProperty(string.Empty);

            // Set by the Settings UI when we want to reinitialize the logic to connect to another PC.
            PendingConnectionRequest = null;
            PendingKeyGenerationRequest = null;

            // TODO(yuyoyuppe): edit hotkey from settings page
            HotkeySwitchMachine = new IntProperty(0x70); // VK.F1

            // These are internal, i.e. cannot be edited directly from UI
            MachinePool = ":,:,:,:";
            MatrixOneRow = false;
            MachineID = new IntProperty(0);
            LastX = new IntProperty(0);
            LastY = new IntProperty(0);
            PackageID = new IntProperty(0);
            FirstRun = true;
            TCPPort = new IntProperty(15100);
            DrawMouseEx = true;
            Name2IP = new StringProperty(string.Empty);
            UseVKMap = false;
            FisrtCtrlShiftS = false;
            StealFocusWhenSwitchingMachine = false;
        }
    }
}
