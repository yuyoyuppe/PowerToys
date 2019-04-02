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
using Accessibility;
using System.Diagnostics;

namespace VerificationRoutines
{
    [Verification(
        "CheckParentChild",
        "TBD",
        Group = ClassificationsGroup.TreeStructure,
        Priority = 1)]
    /// -------------------------------------------------------------------
    /// <summary>Check Parent Child Tests</summary>
    /// <TESTS>
    /// Element is a Child of it's parent
    /// Each of the IAccessible's children's(obtained by AccessibleChildren) accParent property will return the IAccessible
    /// </TESTS>
    /// -------------------------------------------------------------------
    public class CheckParentChild : MsaaTraversingVerificationBase, IVerificationRoutine
    {
        #region LogMessages

        static LogMessage ElementsChildHasDifferentParent = new LogMessage("Element's child {0} has different parent {1} than element)", "ElementsChildHasDifferentParent", 1);
        static LogMessage NullParent = new LogMessage("A call to accParent returned NULL", "NullParent", 1);
        static LogMessage ElementIsNotChildOfElementsParent = new LogMessage("Element's accParent is ({0}), but original element is not part of the parents.AccessibleChildren()", "ElementIsNotChildOfElementsParent", 1);
        static LogMessage ElementIsChildOfParentMulipleTimes = new LogMessage("Element's accParent is ({0}), but original element is a child {1} times", "ElementIsChildOfParentMulipleTimes", 1);

        #endregion LogMessages

        protected override bool IgnoreInvisibleElements
        {
            get
            {
                return true;
            }
        }

        [TestMethod("Element is a Child of its parent")]
        public void AccNavigation_ElementIsChildOfElementsParent(MsaaElement msaaChild)
        {
            Accessible accParent;

            // Get the element's parent
            try
            {
                // MsaaElement does not call accParent to set its parent, so need to default back
                // to IAccessible's accParent to get the element.
                accParent = msaaChild.Accessible.Parent;
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception))
                {
                    throw;
                }

                TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, msaaChild, "when calling accParent");
                return;
            }

            if (accParent == null)
            {
                //"Elements parent is NULL"
                TestLogger.Log(EventLevel.Error, CheckParentChild.NullParent, this.GetType(), msaaChild);
                return;
            }

            // Get the element's parent's children
            Accessible[] children = null;
            int hr = Win32API.S_OK;
            try
            {
                hr = accParent.Children(out children);
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception))
                {
                    throw;
                }

                TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, msaaChild.Accessible, "when calling AccessibleChildren");
                return;
            }

            int found = 0;
            // See if the original element is a child of the element's parent
            foreach (Accessible child in children)
            {
                if (msaaChild.Compare(child))
                {
                    found++;
                    break;
                }
            }

            if (found == 0)
            {
                TestLogger.Log(EventLevel.Error, ElementIsNotChildOfElementsParent, this.GetType(), msaaChild, accParent.ToString());
            }
            else if (found > 1)
            {
                TestLogger.Log(EventLevel.Error, ElementIsChildOfParentMulipleTimes, this.GetType(), msaaChild, accParent.ToString(), found);
            }

        }

        // --------------------------------------------------------------------
        [TestMethod("Each of the IAccessible's children's(obtained by AccessibleChildren) accParent property will return the IAccessible")]
        public void AccNavigation_AccessibleChildrensParentAreSameIAccessible(MsaaElement parentMsaaElement)
        {
            if (parentMsaaElement.GetRole() == AccRole.Window)
            {
                // An element with a role of window is almost always created by oleacc and has the same 
                // properties as the child it puts under itself.  So this is unlikely to find real errors in these 
                // elements.  These most likely happen in hosting scenarios like Scenic or DUI.
                return;
            }
            
            Accessible accParent = null;
            foreach (MsaaElement childMsaaElement in parentMsaaElement.Children)
            {
                // We ignore invisible children when traversing but the tree still has them there
                if (!childMsaaElement.IsVisible)
                {
                    continue;
                }

                // MsaaElement does not call accParent to set it's parent, so need to default back
                // to IAccessible's accParent to get the element.

                try
                {
                    accParent = childMsaaElement.Accessible.Parent;
                }
                catch (Exception exception)
                {
                    if (VerificationHelper.IsCriticalException(exception))
                    {
                        throw;
                    }
                    TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, null, parentMsaaElement, "when calling element's children's IAccessible.Parent", exception.Message);
                    return;
                }
                if (accParent == null)
                {
                    // this is an error but other verification will get this no need to report it twice
                    continue;
                }

                if (!parentMsaaElement.Compare(accParent))
                {
                    TestLogger.Log(EventLevel.Error, ElementsChildHasDifferentParent, this.GetType(), parentMsaaElement, parentMsaaElement.Accessible.ToString(), accParent.ToString());
                }
            }
        }
    }
}
