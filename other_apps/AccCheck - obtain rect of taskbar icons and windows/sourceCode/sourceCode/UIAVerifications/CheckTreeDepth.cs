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
        "CheckTreeDepthUIA",
        "Checks that the UIA Tree is not excessively deep",
        Group = "UIA Properties",
        Priority = 1)]
    public class CheckTreeDepth : IVerificationRoutine
    {
        private const int ERROR_TREE_DEPTH = 25;
        private const int WARNING_TREE_DEPTH = 10;
        private uint _depth = 0;
        private uint _nodeCount = 0;

        #region LogMessages

        private static LogMessage TreeMightBeCyclic = new LogMessage("Tree is at least {0} levels deep. This is too deep of a tree, or might indicate that the accessibility tree is cyclic, and thus it would appear to be infinite.", "TreeMightBeCyclic", 1);
        private static LogMessage TreeTooDeep = new LogMessage("Tree is {0} levels deep.", "TreeTooDeep", 2);
        private static LogMessage DepthChecked = new LogMessage("Tree is {0} levels deep.", "DepthChecked");

        #endregion LogMessages

        private void TraverseTree(UIA.IUIAutomationElement automationElement, uint level)
        {
            ++_nodeCount;

            if (level > _depth)
            {
                _depth = level;
            }

            // never go deeper than ERROR_TREE_DEPTH, that's a sign of a loop
            if (_depth > ERROR_TREE_DEPTH)
            {
                return;
            }

            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    TraverseTree(children.GetElement(i), level + 1);
                }
            }
        }

        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;
            _nodeCount = 0;
            _depth = 0;

            UIA.IUIAutomationElement root = UIAGlobalContext.CacheUIATree(hwnd);
            TraverseTree(root, 0);

            if (_depth >= ERROR_TREE_DEPTH)
            {
                LogHelper.Log(EventLevel.Error, TreeMightBeCyclic, this.GetType(), UIAGlobalContext.LocalRootElement, _depth);
            }
            else if (_depth >= WARNING_TREE_DEPTH)
            {
                LogHelper.Log(EventLevel.Warning, TreeTooDeep, this.GetType(), UIAGlobalContext.LocalRootElement, _depth);
            }
            else
            {
                LogHelper.Log(EventLevel.Information, DepthChecked, this.GetType(), UIAGlobalContext.LocalRootElement, _depth);
            }

            AccCheck.Logging.UsageReporter.Instance.ReportUIComplexity(
                _nodeCount, _depth);
        }
    }
}
