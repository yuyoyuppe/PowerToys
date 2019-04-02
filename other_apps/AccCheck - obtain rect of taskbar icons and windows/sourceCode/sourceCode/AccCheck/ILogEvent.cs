// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace AccCheck.Logging
{
    [ComVisible(true), GuidAttribute("7DF210F1-6E00-4ae5-AD80-21AA49FACA2C")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ILogEvent
    {
        EventLevel Level
        {
            get;
        }
        String EventID
        {
            get;
        }
        String ExtraInformation
        {
            get;
        }
        String Text
        {
            get;
        }
        String ParentChain
        {
            get;
        }
        String VerificationRoutine
        {
            get;
        }
        String Classname
        {
            get;
        }
        String AccName
        {
            get;
        }
        String AccValue
        {
            get;
        }
        string AccRole
        {
            get;
        }
        string AccState
        {
            get;
        }
        DateTime Timestamp
        {
            get;
        }
        int SequenceNumber
        {
            get;
        }
        bool Suppressed
        {
            get;
        }
    }
}
