// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using AccCheck.Logging;

namespace AccCheck.Verification
{
    /*
     * VerificationRoutines gets a handle to a window, a logger, and a boolean
     * that tells if the app is running graphically.
     * You need to override the Execute method, and specify a metadata attribute
     */
    public abstract class VerificationRoutine
    {
        public abstract void Execute(IntPtr hwnd, Logger logger, bool AllowUI);
    }
}
