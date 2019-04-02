// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;

namespace VerificationRoutines
{
    [Verification(
        "CheckName",
        "TBD",
        Group = ClassificationsGroup.Properties,
        Priority = 1)]
    /// -------------------------------------------------------------------
    /// <summary>Check accName Tests</summary>
    /// <TESTS>
    /// Test : Verify IAccessible that is focusable has a valid accName.
    /// Test : The control's accName does not start with the text of the accRole.
    /// Test : Name is shorter than defined length.
    /// Test : Verify IAcessible's accName does not include invalid characters"]
    /// </TESTS>
    /// -------------------------------------------------------------------
    public class CheckName : MsaaTraversingVerificationBase, IVerificationRoutine
    {
        const int CHAR_DOES_NOT_EXISTS = -1;

        #region LogMessages

        static LogMessage ElementHasNoName = new LogMessage("Element has no name", "ElementHasNoName", 1);
        static LogMessage AccNameLengthTooLong = new LogMessage("Element's accName should not return a string longer than {0} characters", "AccNameLengthTooLong", 2);
        static LogMessage AccNameShouldNotContainRole = new LogMessage("accName \"{0}\" should not contain the text of the Role \"{1}\"", "AccNameShouldNotContainRole", 2);

        #endregion LogMessages

        private MsaaElement [] _firstNoNameErrors = new MsaaElement[4];
        private AccRole [] _suppressFirstErrorRoles = new AccRole [] {AccRole.List, AccRole.Tree, AccRole.Toolbar, AccRole.Dialog};

        protected override bool IgnoreInvisibleElements
        {
            get
            {
                return true;
            }
        }
        
        protected override void Initialize()
        {
            for (int i = 0; i < _firstNoNameErrors.Length; i++)
            {
                _firstNoNameErrors[i] = null;
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Verify IAccessible that is focusable has a valid accName"</summary>
        /// <TODO>Currently nothing returns STATE_SYSTEM_FOCUSABLE</TODO>
        /// -------------------------------------------------------------------
        [TestMethod("Verify element that is focusable has a valid accName")]
        public void AccName_FocusControlsHaveNames(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            if (element.IsStatePresent(AccState.Focusable) && !element.IsStatePresent(AccState.Unavailable))
            {
                // Title bar does not have a name but does have a value
                AccRole role = element.GetRole();
                if (role != AccRole.Titlebar)
                {
                    string name = element.GetName();
                    if (string.IsNullOrEmpty(name))
                    {
                        bool errorSuppressed = false;
                        
                        for (int i = 0; i < _firstNoNameErrors.Length; i++)
                        {
                            if  (_suppressFirstErrorRoles[i] == role) 
                            {   
                                if (_firstNoNameErrors[i] == null)
                                {
                                    errorSuppressed = true;
                                    _firstNoNameErrors[i] = element;
                                }
                                else
                                {
                                    TestLogger.Log(EventLevel.Error, ElementHasNoName, this.GetType(), _firstNoNameErrors[i]);
                                    // put a bad value in here so from now we just log all no names for this control type
                                    _suppressFirstErrorRoles[i] = AccRole.Invalid;
                                }
                                break;
                            }
                        }
                        
                        if (!errorSuppressed)
                        {
                            //"There is no name for the element");
                            TestLogger.Log(EventLevel.Error, ElementHasNoName, this.GetType(), element);
                        }
                    }
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>The control's accName does not contain text of the 
        /// accRole (e.g. no Toolbar accName of 'Toolbar5')</summary>
        /// -------------------------------------------------------------------
        [TestMethod("The control's accName does not contain the text of the accRole")]
        public void AccName_FocusableControlsNameDoesNotIncludeRoleText(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            string name = element.GetName();
            
            if (name != null && element.IsStatePresent(AccState.Focusable) && !element.IsStatePresent(AccState.Unavailable))
            {
                string lname = name.ToLower();

                // There is too many places where 'window' is used such as Microsoft Windows to do 
                // a valid check so bypass validation check for this.
                if (!lname.Contains("window"))
                {
                    {
                        AccRole role = element.GetRole();
                        string lrole = role.ToString().ToLower();

                        // People often use the role in the name correctly like "Explorer pane"
                        // but starting with the role is usually bad like "pane1"
                        // If the name has a space in it it is probably not something that defaulted like "button_1"
                        if (lname.StartsWith(lrole) && !lname.Contains(" "))
                        {
                            // "accName ({0}) should not contain the text of the Role ({1})"
                            TestLogger.Log(EventLevel.Error, AccNameShouldNotContainRole, this.GetType(), element, name, role);
                        }
                    }
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Name is shorter than defined length (32000 chars)</summary>
        /// -------------------------------------------------------------------
        [TestMethod("Name is shorter than defined length")]
        public void AccName_NameStringNotTooLong(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            const int MAX_NAME_LENGTH = 32000;

            string name = element.GetName();

            if (name != null)
            {
                if (name.Length > MAX_NAME_LENGTH)
                {
                    TestLogger.Log(EventLevel.Error, AccNameLengthTooLong, this.GetType(), element, MAX_NAME_LENGTH);
                }
            }
        }
    }
}
