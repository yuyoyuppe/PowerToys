// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

//-----------------------------------------------------------------------------
// Description: This is the set of tests that are used within the AccChecker application
//-----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Text;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;

namespace VerificationRoutines
{
    [Verification(
        "CheckShortCuts",
        "TBD",
        false,
        ClassificationsGroup.Properties)]
    /// -------------------------------------------------------------------
    /// <summary>Check Shortcut Tests</summary>
    /// <TESTS></TESTS>
    /// -------------------------------------------------------------------
    public class CheckShortCuts : VerificationRoutinesBase, IVerificationRoutine
    {
        public CheckShortCuts()
            : base(typeof(CheckShortCuts))
        {
            this.VerificationRoutineThisObject = this;
        }
        
        #region TestSteps
        #endregion TestSteps

    }
}
