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
using UIA=UIAutomationClient;

namespace UIAVerifications
{
    [Verification(
        "CheckNavigateUia",
        "Checks that UIA TreeWalker Navigation functions consistently throughout the tree",
        Group = "UIA Tree",
        Priority = 1)]
    public class CheckNavigate : IVerificationRoutine
    {
        enum TreeWalkerMethod
        {
            GetNextSiblingElement,
            GetPreviousSiblingElement,
            GetParentElement,
            GetFirstChildElement,
            GetLastChildElement,
        };

        private const int MAX_CIRCULAR_REFERENCE = 50;
        private const int ERROR_TREE_DEPTH = 25;
        private uint _depth = 0;
 
        #region LogMessages

        static LogMessage TooManyChildren = new LogMessage("Stopped looking for additional child siblings after finding {0} siblings.  Possible incorrect tree", "TooManyChildren", 2);
        static LogMessage IncorrectRoundTrip = new LogMessage("Incorrect roundtrip navigation : {0}", "IncorrectRoundTrip", 1);
        static LogMessage Navigation_ReturnedElementWithIncorrectParent = new LogMessage("When navigating by calling {0} from the tested node, and then repeatedly calling {1}, the following child was found: {2}. The child reported its parent is {3}. Instead, the child should report as its parent the testing node.", "Navigation_ReturnedElementWithIncorrectParent", 1);
        static LogMessage FirstLastMismatch = new LogMessage("When navigating by calling {0} from the tested node, and then repeatedly calling {1}, we finished at the following child: {2}. This did not match the result of calling {3} on the tested node, which yielded: {4}. Instead, the child should report as its parent the testing node.", "FirstChildLastChildMismatch", 1);

        #endregion LogMessages

        /// -------------------------------------------------------------------
        /// <summary>Element = all children's accParent should return you back the the original Element</summary>
        /// -------------------------------------------------------------------
        void Test_ChildsParents(UIA.IUIAutomationElement parentElement, UIA.IUIAutomationTreeWalker walker, TreeWalkerMethod downDirection, TreeWalkerMethod downDirection2, TreeWalkerMethod traverseDirection)
        {
            // Navigate down, then navigate across siblings and verify that they are all children 
            // of the original element.
            UIA.IUIAutomationElement nextElement = null;

            // Navigate down
            try
            {
                nextElement = Navigate(parentElement, walker, downDirection);
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception))
                {
                    throw;
                }

                LogHelper.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, parentElement, string.Format("TreeWalker.{0})", downDirection));
                return;
            }

            if (nextElement == null)
            {
                return;
            }

            int loopCount = 0;
            UIA.IUIAutomationElement nextElementParent = null;

            // Now navigate all siblings and verify that all of their parents are the original parent passed into the method.
            for (; ; loopCount++)
            {
                try
                {
                    nextElementParent = walker.GetParentElement(nextElement);
                }
                catch (Exception exception)
                {
                    if (VerificationHelper.IsCriticalException(exception))
                    {
                        throw;
                    }

                    LogHelper.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, parentElement, "TreeWalker.GetParentElement");
                }

                // Compare that the parent is correct.
                if (nextElementParent == null || UIAGlobalContext.Automation.CompareElements(parentElement, nextElementParent) == 0)
                {

                    LogHelper.Log(EventLevel.Error, Navigation_ReturnedElementWithIncorrectParent, this.GetType(), parentElement,
                        downDirection, traverseDirection, SummaryString(nextElement), SummaryString(nextElementParent));
                    return;
                }

                // Make sure we don't get into a circular situation.
                if (loopCount > MAX_CIRCULAR_REFERENCE)
                {
                    // "Stopped looking for additional child siblings after finding {0} siblings.  Possible incorrect tree"
                    LogHelper.Log(EventLevel.Warning, TooManyChildren, this.GetType(), parentElement, MAX_CIRCULAR_REFERENCE);
                    return;
                }

                // Navigate to the next sibling.
                try
                {
                    UIA.IUIAutomationElement nextElementTemp = Navigate(nextElement, walker, traverseDirection);
                    if (nextElementTemp == null)
                    {
                        // We've navigated to either a sibling that is broken and not pointing to the correct next sibling, 
                        // or we have gone through them all.  At this point, check the send down navigation method and make
                        // sure it matches.
                        UIA.IUIAutomationElement lastElement = null;

                        // Navigate down
                        lastElement = Navigate(parentElement, walker, downDirection2);

                        // Compare that the parent is correct.
                        if (lastElement == null || UIAGlobalContext.Automation.CompareElements(nextElement, lastElement) == 0)
                        {

                            LogHelper.Log(EventLevel.Error, FirstLastMismatch, this.GetType(), parentElement,
                                downDirection, traverseDirection, SummaryString(nextElement), downDirection2, SummaryString(lastElement));
                            return;
                        }

                        // If it all checks out, return
                        return;
                    }
                    nextElement = nextElementTemp;
                }
                catch (Exception exception)
                {
                    if (VerificationHelper.IsCriticalException(exception))
                    {
                        throw;
                    }

                    LogHelper.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, nextElement, "Test_ChildsParent");
                    return;
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Navigate one direction, and then navigate back, then the 1st direction...should be the 1st object we navigated to</summary>
        /// -------------------------------------------------------------------
        void Test_UiaNavigateRoundTrip(UIA.IUIAutomationElement automationElement, UIA.IUIAutomationTreeWalker walker, TreeWalkerMethod direction1, TreeWalkerMethod direction2)
        {
            UIA.IUIAutomationElement element0 = automationElement;

            UIA.IUIAutomationElement elementTemp = null;
            UIA.IUIAutomationElement element1 = null;
            UIA.IUIAutomationElement element2 = null;
            UIA.IUIAutomationElement element3 = null;

            try
            {
                element1 = Navigate(element0, walker, direction1);
                if (element1 == null)
                {
                    // Simply unable to navigate in that direction
                    return;
                }

                element2 = Navigate(element1, walker, direction2);
                if (element2 == null || (UIAGlobalContext.Automation.CompareElements(element0, element2) == 0))
                {
                    string buffer = string.Format("Call to TreeWalker.{3}({0}) returned {1},\nthen TreeWalker.{4}({1}) returned {2} and should have returned {0}",
                        SummaryString(element0),
                        SummaryString(element1),
                        SummaryString(element2),
                        direction1, direction2);

                    //"Incorrect roundtrip navigation : {0}"
                    LogHelper.Log(EventLevel.Error, IncorrectRoundTrip, this.GetType(), automationElement, buffer);
                    return;
                }

                // This should never return NULL element, we've already done this direction above.
                element3 = Navigate(element2, walker, direction1);
                if (element3 == null || (UIAGlobalContext.Automation.CompareElements(element1, element3) == 0))
                {
                    string buffer = string.Format("Call to TreeWalker.{4}({0}) returned {1},\nthen TreeWalker.{5}({1}) returned {2},\nthen TreeWalker.{4}({2}) returned {3} and should have returned {1}",
                        SummaryString(element0),
                        SummaryString(element1),
                        SummaryString(element2),
                        SummaryString(element3), direction1, direction2);

                    //"Incorrect roundtrip navigation : {0}"
                    LogHelper.Log(EventLevel.Error, IncorrectRoundTrip, this.GetType(), automationElement, buffer);
                    return;
                }

            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception))
                {
                    throw;
                }

                LogHelper.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, elementTemp, "Test_UiaNavigateRoundTrip");
                return;
            }
        }

        UIA.IUIAutomationElement Navigate(UIA.IUIAutomationElement automationElement, UIA.IUIAutomationTreeWalker walker, TreeWalkerMethod method)
        {
            switch (method)
            {
                case TreeWalkerMethod.GetNextSiblingElement:
                    return walker.GetNextSiblingElement(automationElement);
                case TreeWalkerMethod.GetPreviousSiblingElement:
                    return walker.GetPreviousSiblingElement(automationElement);
                case TreeWalkerMethod.GetParentElement:
                    return walker.GetParentElement(automationElement);
                case TreeWalkerMethod.GetFirstChildElement:
                    return walker.GetFirstChildElement(automationElement);
                case TreeWalkerMethod.GetLastChildElement:
                    return walker.GetLastChildElement(automationElement);
                default:
                    return null;
            }
        }

        string SummaryString(UIA.IUIAutomationElement automationElement)
        {
            if (automationElement == null)
            {
                return "NULL";
            }
            string name = automationElement.CurrentName;
            string controlType = automationElement.CurrentLocalizedControlType;
            name = String.IsNullOrEmpty(name) ? "(nothing)" : name;
            return String.Format("{0}[{1}]", name, controlType);
        }

        private void TraverseTree(UIA.IUIAutomationElement automationElement, UIA.IUIAutomationTreeWalker walker, uint level)
        {
            if (level > _depth)
            {
                _depth = level;
            }

            // never go deeper than ERROR_TREE_DEPTH, that's a sign of a loop
            if (_depth > ERROR_TREE_DEPTH)
            {
                return;
            }

            Test_UiaNavigateRoundTrip(automationElement, walker, TreeWalkerMethod.GetNextSiblingElement, TreeWalkerMethod.GetPreviousSiblingElement);
            Test_UiaNavigateRoundTrip(automationElement, walker, TreeWalkerMethod.GetPreviousSiblingElement, TreeWalkerMethod.GetNextSiblingElement);
            Test_ChildsParents(automationElement, walker, TreeWalkerMethod.GetFirstChildElement, TreeWalkerMethod.GetLastChildElement, TreeWalkerMethod.GetNextSiblingElement);
            Test_ChildsParents(automationElement, walker, TreeWalkerMethod.GetLastChildElement, TreeWalkerMethod.GetFirstChildElement, TreeWalkerMethod.GetPreviousSiblingElement);

            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    TraverseTree(children.GetElement(i), walker, level + 1);
                }
            }
        }

        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;
 
            UIA.IUIAutomationTreeWalker rawWalker = UIAGlobalContext.Automation.RawViewWalker;
            UIA.IUIAutomationElement root = UIAGlobalContext.CacheUIATree(hwnd);
            _depth = 0;
            TraverseTree(root, rawWalker, 0);
        }
    }
}
