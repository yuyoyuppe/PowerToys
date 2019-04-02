// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using AccCheck.Logging;
using AccCheck.Verification;
using System.Runtime.InteropServices;

namespace AccCheck.Verification
{
    [ComVisible(true), GuidAttribute("2E71A815-CE3A-491C-B639-DAD71ACB89F3")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IVerificationRoutineData
    {
        string Id {get; set;}
        bool Active {get; set;}
        String Title { get; }
        String Description { get; }
        bool NeedsUI { get; }
        String Group { get; }
        bool CanVisualize { get; } 
        Type Type { get; }
    }
}