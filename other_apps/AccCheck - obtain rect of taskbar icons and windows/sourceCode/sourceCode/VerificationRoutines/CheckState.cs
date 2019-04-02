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
        "CheckState",
        "Checks the state property of MSAA",
        Group = ClassificationsGroup.Properties,
        Priority = 2)]
    public class CheckState : MsaaTraversingVerificationBase, IVerificationRoutine
    {
        private static LogMessage ElementShouldBeOffScreen = new LogMessage("Element should be offscreen", "ElementShouldBeOffScreen", 2);
        private static LogMessage InconsistentState = new LogMessage("A control has states that are inconsistent {0} and {1}", "InconsistentState", 2);

        protected override bool IgnoreInvisibleElements
        {
            get
            {
                return true;
            }
        }

        //
        // Verifiy the the variant we get back from get_accState is valid.
        // If there was an exception that occured when the method was 
        // call report it if it is an error.
        [TestMethod("Verifiy the the variant we get back from get_accState is valid")]
        public void AccState_CheckVariant(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            Exception exception;
            AccState [] state = element.GetStates(out exception);

            if (exception == null)
            {
                // nothing to report
                return;
            }
            else if (exception is VariantNotIntException)
            {
                VariantNotIntException variantNotIntException = (VariantNotIntException)exception;
                TestLogger.Log(EventLevel.Error, LogEvent.VariantNotInt, this.GetType(), element, "get_accState", "int", variantNotIntException.VariantType);
            }
            else 
            {
                try
                {
                    // This is a exception that was cached (probably when we were constructing the tree).  
                    // Wrap the exception so we get the stack trace to this method, with the inner exception to 
                    // the stack trace to the original call.  Using 'PREVIOUSTHROW' will tell the LogException
                    // how to handle this special case.
                    throw new Exception("PREVIOUSTHROWget_accState()", exception);
                }
                catch(Exception outerException)
                {
                    TestLogger.LogException(EventLevel.Warning, LogEvent.MethodExceptionOccured, this.GetType(), outerException, null, element, exception.GetType().ToString());
                }
            }
        }

        //
        // There are some state that just don't make any sense
        // Find these here and report them as errors.
        //
        [TestMethod("Verify any inconsistencies with the value returned from accState")]
        public void AccState_Inconsistent(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }
            
            AccState[] states = element.GetStates();

            bool containsCollapsed = Array.IndexOf(states, AccState.Collapsed) != -1;
            bool containsExpanded = Array.IndexOf(states, AccState.Expanded) != -1;

            // Something can't be expanded and collaped
            if (containsCollapsed && containsExpanded)
            {
                TestLogger.Log(EventLevel.Error, InconsistentState, this.GetType(), element, AccState.Collapsed, AccState.Expanded);
            }

            bool containsSelected = Array.IndexOf(states, AccState.Selected) != -1;
            bool containsSelectable = Array.IndexOf(states, AccState.Selectable) != -1;

            // If something is selected it must be selectable
            if (containsSelected && !containsSelectable)
            {
                TestLogger.Log(EventLevel.Error, InconsistentState, this.GetType(), element, AccState.Selected, "not " + AccState.Selectable);
            }

            bool containsFocused = Array.IndexOf(states, AccState.Focused) != -1;
            bool containsFocusable = Array.IndexOf(states, AccState.Focusable) != -1;

            // If something is focused it must be focusable
            if (containsFocused && !containsFocusable)
            {
                if (element.GetRole() != AccRole.MenuItem && element.Children.Count == 0)
                {
                    TestLogger.Log(EventLevel.Error, InconsistentState, this.GetType(), element, AccState.Focused, "not " + AccState.Focusable);
                }
            }
        }

                    
        [TestMethod("Verify the offscreen state is correct")]
        public void AccState_ShouldBeOffScreen(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }

            // We only wnat to do this verification on leaf nodes
            if (element.Children.Count > 0)
            {
                return;
            }

            Rectangle rect = element.GetBoundingBox();
            if (rect == Rectangle.Empty || rect.Width == 0 || rect.Height == 0)
            {
                return;
            }

            // If the rect is clipped by the parent there is nothing left it should be offscreen
            MsaaElement parent = element.Parent;
            Rectangle rectParent = parent.GetBoundingBox();
            rect.Intersect(rectParent);

            if (rect.IsEmpty)
            {
                if (!element.IsStatePresent(AccState.OffScreen))
                {
                    TestLogger.Log(EventLevel.Warning, ElementShouldBeOffScreen, this.GetType(), element);
                }

                return;
            }


        }

    }
}
