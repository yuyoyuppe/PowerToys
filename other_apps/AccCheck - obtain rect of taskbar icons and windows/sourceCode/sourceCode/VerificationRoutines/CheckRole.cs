// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using System.Drawing;
using System.Reflection;

namespace VerificationRoutines
{
    [Verification(
        "CheckRole",
        "Checks the role property of MSAA",
        Group = ClassificationsGroup.Properties,
        Priority = 1)]
    public class CheckRole : MsaaTraversingVerificationBase, IVerificationRoutine
    {
        #region LogMessages

        static LogMessage InvalidRole = new LogMessage("The int returned from get_accRole is not a valid role constant, ie ROLE_SYSTEM_*", "InvalidRole", 1);
        static LogMessage ControlShouldHaveValue = new LogMessage("A control with role of {0} should have a value but get_accValue is not implemented", "ControlShouldHaveValue", 1);
        static LogMessage MemberNotImplememented = new LogMessage("The method is not implemented on this particular element.", "MemberNotImplememented", 3);

        #endregion LogMessages

        protected override bool IgnoreInvisibleElements
        {
            get
            {
                return true;
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        /// 1. Verifiy the the variant we get back from get_accRole is valid.
        ///    If there was an exception that occured when the method was 
        ///    call report it if it is an error.
        /// </summary>
        /// -------------------------------------------------------------------
        [TestMethod]
        public void AccRole_CheckVariant(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            Exception exception;
            AccRole role = element.GetRole(out exception);
            if (exception == null)
            {
                if (role == AccRole.Invalid)
                {
                    TestLogger.Log(EventLevel.Error, InvalidRole, this.GetType(), element, null);
                }
            }
            else if (exception is VariantNotIntException)
            {
                VariantNotIntException variantNotIntException = (VariantNotIntException)exception;
                TestLogger.Log(EventLevel.Error, LogEvent.VariantNotInt, this.GetType(), element, "get_accRole", "int", variantNotIntException.VariantType);
            }
            else 
            {
                try
                {
                    // This is a exception that was cached (probably when we were constructing the tree).  
                    // Wrap the exception so we get the stack trace to this method, with the inner exception to 
                    // the stack trace to the original call.  Using 'PREVIOUSTHROW' will tell the LogException
                    // how to handle this special case.
                    //get_accRole()
                    throw new Exception("PREVIOUSTHROWget_accRole()", exception);
                }
                catch (Exception outerException)
                {
                    TestLogger.LogException(EventLevel.Warning, LogEvent.MethodExceptionOccured, this.GetType(), outerException, null, element, exception.GetType().ToString());
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        /// 2. There are some roles that need to have a value associated with them.
        /// The role here should all support value if they don't it's an error
        /// </summary>
        /// -------------------------------------------------------------------
        [TestMethod]
        public void AccRole_NeedsValue(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            AccRole role = element.GetRole();
            if (role == AccRole.Combobox ||
                 role == AccRole.Slider ||
                 role == AccRole.Text ||
                 role == AccRole.SpinButton ||
                 role == AccRole.Scrollbar ||
                 role == AccRole.ProgressBar ||
                 role == AccRole.TreeItem ||
                 role == AccRole.IPAddress)
            {
                Exception exception;
                string value = element.GetValue(out exception);
                if (exception is NotImplementedException)
                {
                    TestLogger.Log(EventLevel.Error, ControlShouldHaveValue, this.GetType(), element, role.ToString());
                }
            }
        }
    }
}
