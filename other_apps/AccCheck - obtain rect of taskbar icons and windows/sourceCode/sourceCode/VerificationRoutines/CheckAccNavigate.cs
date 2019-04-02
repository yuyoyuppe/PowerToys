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
using Accessibility;
using System.Diagnostics;

namespace VerificationRoutines
{
    [Verification(
        "CheckNavigate",
        "TBD",
        Group = ClassificationsGroup.TreeStructure)]
    /// -------------------------------------------------------------------
    /// <summary>Check accNavigate Tests</summary>
    /// <TESTS>
    ///     1. From element, navigate using NAVDIR_NEXT, NAVDIR_PREV should return you back the same element (
    ///     1. From element, navigate using NAVDIR_PREV, NAVDIR_NEXT should return you back the same element 
    /// 
    ///     2. From element, navigate using Win32API.NAVDIR_FIRSTCHILD, then navigate all siblings' using NAVDIR_NEXT and verify that accParent returns back to the original element
    ///     3. From element, navigate using Win32API.NAVDIR_LASTCHILD, then navigate all siblings' using NAVDIR_NEXT and verify that accParent returns back to the original element
    /// </TESTS>
    /// -------------------------------------------------------------------
    public class CheckNavigate : MsaaTraversingVerificationBase, IVerificationRoutine
    {
        const int MAX_CIRCULAR_REFERENCE = 50;

        #region LogMessages

        static LogMessage TooManyChildren = new LogMessage("Stopped looking for additional child siblings after finding {0} siblings.  Possible incorrect tree", "TooManyChildren");
        static LogMessage IncorrectRoundTrip = new LogMessage("Incorrect roundtrip navigation : {0}", "IncorrectRoundTrip");
        static LogMessage AccNavigate_ReturnedElementWithIncorrectParent = new LogMessage("Element's child has a different parent than element", "AccNavigate_ReturnedElementWithIncorrectParent");

        #endregion LogMessages

        protected override bool IgnoreInvisibleElements
        {
            get
            {
                return true;
            }
        }

        // Valid round trip navigation for tests
        NavDir[] _directions =  
                    { 
                        // Covered in accParent tests: Win32API.NavDir.NAVDIR_FIRSTCHILD, Win32API.NavDir.NAVDIR_LASTCHILD, 
                        // Spacial scenario: Win32API.NavDir.NAVDIR_DOWN, Win32API.NavDir.NAVDIR_UP, 
                        // Spacial scenario: Win32API.NavDir.NAVDIR_UP, Win32API.NavDir.NAVDIR_DOWN, 
                        // Spacial scenario: Win32API.NavDir.NAVDIR_LEFT, Win32API.NavDir.NAVDIR_RIGHT, 
                        // Spacial scenario: Win32API.NavDir.NAVDIR_RIGHT, Win32API.NavDir.NAVDIR_LEFT, 
                        NavDir.next, NavDir.previous, 
                        NavDir.previous, NavDir.next
                    };

        // -------------------------------------------------------------------
        [TestMethod(Description = "Round trip navigating using accNavigate (ie. NAVDIR_NEXT, NAVDIR_PREV) returns the original IAccessible")]
        public void AccNavigation_RoundTrip(MsaaElement element)
        {
            for (int i = 0; i < _directions.GetLength(0); )
            {
                Test_accNavigateRoundTrip(element, _directions[i], _directions[i + 1]);
                i++; i++; // _directions is really a pair in the array
            }
        }

        // -------------------------------------------------------------------
        [TestMethod(Description = "From an IAccessible, accNavigate(NAVDIR_FIRSTCHILD) and then all siblings using accNavigate(NAVDIR_NEXT) and verify all elements accParent is the original element")]
        public void AccNagivate_ChildrensParents_NAVDIR_FIRST_NAVDIR_NEXT(MsaaElement element)
        {
            Test_ChildsParents(element, NavDir.firstChild, NavDir.next);
        }

        // -------------------------------------------------------------------
        [TestMethod(Description = "From an IAccessible, accNavigate(NAVDIR_LASTCHILD) and then all siblings using accNavigate(NAVDIR_PREV) and verify all elements accParent is the original element")]
        public void AccNagivate_ChildrensParents_NAVDIR_LAST_NAVDIR_PREV(MsaaElement element)
        {
            Test_ChildsParents(element, NavDir.lastChild, NavDir.previous);
        }

        #region TestSteps

        /// -------------------------------------------------------------------
        /// <summary>Element = all children's accParent should return you back the the original Element</summary>
        /// -------------------------------------------------------------------
        void Test_ChildsParents(MsaaElement parent, NavDir downDirection, NavDir traverseDirection)
        {
            // Navigate down, then navigate across siblings and verify that they are all children 
            // of the original element.
            Accessible nextElement = null;

            // Exit out if we can't navigate down
            if (parent.Children.Count < 1)
            {
                return;
            }

            // Navigate down
            try
            {
                nextElement = parent.Accessible.Navigate(downDirection);
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception)) 
                {
                    throw; 
                }
                
                TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, parent.Accessible, string.Format("accNavigate({0})", downDirection));
                return;
            }

            if (nextElement == null)
            {
                return;
            }

            int loopCount = 0;
            Accessible nextElementParent = null;

            // Now navigate all siblings and verify that all of their parents are the original parent passed into the method.
            for (; ; loopCount++)
            {
                try
                {
                    nextElementParent = nextElement.Parent;
                }
                catch (Exception exception)
                {
                    if (VerificationHelper.IsCriticalException(exception)) 
                    {
                        throw; 
                    }
                    
                    TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, parent, "accParent");
                }

                // Compare that the parent is correct.
                if (!parent.Compare(nextElementParent))
                {
                    TestLogger.Log(EventLevel.Error, AccNavigate_ReturnedElementWithIncorrectParent, this.GetType(), parent);
                    return;
                }

                // Make sure we don't get into a circular situation.
                if (loopCount > MAX_CIRCULAR_REFERENCE)
                {
                    // "Stopped looking for additional child siblings after finding {0} siblings.  Possible incorrect tree"
                    TestLogger.Log(EventLevel.Warning, TooManyChildren, this.GetType(), parent, MAX_CIRCULAR_REFERENCE);
                    return;
                }

                // Navigate to the next sibling.
                try
                {
                    Accessible nextElementTemp = nextElement.Navigate(traverseDirection);
                    if (nextElementTemp == null)
                    {
                        // We've navigated to either a sibling that is broken and not pointing to the correct next sibling, 
                        // or we have gone through them all.  Either case, nothing more do do, so exit out.
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
                    
                    TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, nextElement, "Test_ChildsParent");
                    return;
                }
            } 
        }

        /// -------------------------------------------------------------------
        /// <summary>Navigate one direction, and then navigate back, then the 1st direction...should be the 1st object we navigated to</summary>
        /// -------------------------------------------------------------------
        void Test_accNavigateRoundTrip(MsaaElement element, NavDir direction1, NavDir direction2)
        {
            /* AccNavigate will navigate to the next visible element, or next element if no siblings are visible.
             * 
             * Test: Navigate to the next element in one direction(IAcc will get you to next visible element if there is one), 
             * then to the next element in the other direction and then back in the original direction.  Should always get to 
             * the element resulting from the 1st accNavigate (element1).
             * 
             * Root
             *    Element1 (Hidden)
             *    Element2 (Visible)             
             *    Element3 (Hidden)
             *          Child1 (Hidden)
             *          Child2 (Visible)
             *          Child3 (Hidden)
             *          Child4 (Hidden)
             *    Element4 (Hidden)
             *    Element5 (Visible)
             *    Element6 (Hidden)
             * 
             * Scenario w/Element2 (Previous, Next, Previous)
             *      Null...exit w/pass
             * 
             * Scenario w/Element4 (Previous, Next, Previous)
             *      Element2
             *      Element5
             *      Element2...exit w/pass if Element2, else fail
             * 
             * Scenario w/Element5 (Previous, Next, Previous)
             *      Element2
             *      Element5
             *      Element2...exit w/pass if Element2, else fail
             * 
             * Scenario w/Child3 (NOTE: ChildId's seem to always navigate to another child unless there is an IAccessible
             *      Child2
             *      Child3
             *      Child2...exit w/pass if Child2, else fail
             */
            Accessible element0 = element.Accessible;

            Accessible elementTemp = null;
            Accessible element1 = null;
            Accessible element2 = null;
            Accessible element3 = null;

            try
            {
                // MSAA.Current.GetAccNavigate returns false if element1 is null
                element1 = element0.Navigate(direction1);
                if (element1 == null)
                {
                    return;
                }

                element2 = element1.Navigate(direction2);
                if (element2 == null)
                {
                    return;
                }

                // This should never return NULL element, we've already done this direction above.
                element3 = element2.Navigate(direction1);
                if (element3 == null)
                {
                    string buffer = string.Format("Call to {0}.accNavigate({4}) returned {1},\nthen {1}.accNavigate({5}) returned {2},\nthen {2}.accNavigate({4}) returned {3} and should have returned {1}",
                        element0, 
                        element1, 
                        element2, 
                        "NULL", direction1, direction2);

                    //"Incorrect roundtrip navigation : {0}"
                    TestLogger.Log(EventLevel.Error, IncorrectRoundTrip, this.GetType(), element, buffer);
                    return;
                }

            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception)) 
                {
                    throw; 
                }
                
                TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, VerificationHelper.AcceptableCommonErrorCodes, elementTemp, "Test_accNavigateRoundTrip");
                return;
            }

            if (!element1.Compare(element3))
            {
                string buffer = string.Format("Call to {0}.accNavigate({4}) returned {1},\nthen {1}.accNavigate({5}) returned {2},\nthen {2}.accNavigate({4}) returned {3} and should have returned {1}",
                        element0, 
                        element1,
                        element2, 
                        element3, 
                        direction1, direction2);

                //"Incorrect roundtrip navigation : {0}"
                TestLogger.Log(EventLevel.Error, IncorrectRoundTrip, this.GetType(), element, new object[] { buffer });
            }
        }
    }
    #endregion TestSteps
}
