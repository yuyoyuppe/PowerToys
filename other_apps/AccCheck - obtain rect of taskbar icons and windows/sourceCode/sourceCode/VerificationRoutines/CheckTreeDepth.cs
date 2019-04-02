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
    /*
     * Simple tree depth check. Logs an error if the tree is more
     * than 250 levels deep, and gives a warning if it is more than
     * 15 levels deep. Otherwise it logs a message giving the depth
     * of the tree.
     */

    [Verification(
        "CheckTreeDepth",
        "Checks that the accessibility tree isn't excessively deep.",
        Group = ClassificationsGroup.Navigation,
        Priority = 1)]
    public class CheckTreeDepth : IVerificationRoutine
    {
        private const int ERROR_TREE_DEPTH = 50;
        private const int WARNING_TREE_DEPTH = 20;
        private uint _depth = 0;
        private uint _nodeCount = 0;

        #region LogMessages

        private static LogMessage TreeMightBeCyclic = new LogMessage("Tree is {0} levels deep. This might indicate that the accessibility tree is cyclic, and thus it would appear to be infinite.", "TreeMightBeCyclic", 1);
        private static LogMessage TreeTooDeep = new LogMessage("Tree is {0} levels deep.", "TreeTooDeep", 2);
        private static LogMessage DepthChecked = new LogMessage("Tree is {0} levels deep.", "DepthChecked");

        #endregion LogMessages

        private void TraverseTree(MsaaElement parent, uint level)
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

            foreach (MsaaElement child in parent.Children)
            {
                TraverseTree(child, level + 1);
            }
        }

        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            TestLogger.Logger = logger;
            TestLogger.RootHwnd = hwnd;
            _nodeCount = 0;
            _depth = 0;
            
            MsaaElement root = MsaaElement.CacheMsaaTree(hwnd);
            TraverseTree(root, 0);

            if (_depth >= ERROR_TREE_DEPTH)
            {
                TestLogger.Log(EventLevel.Error, TreeMightBeCyclic, this.GetType(), MsaaElement.RootElement, _depth);
            }
            else if (_depth >= WARNING_TREE_DEPTH)
            {
                TestLogger.Log(EventLevel.Warning, TreeTooDeep, this.GetType(), MsaaElement.RootElement, _depth);
            }
            else
            {
                TestLogger.Log(EventLevel.Information, DepthChecked, this.GetType(), MsaaElement.RootElement, _depth);
            }

            AccCheck.Logging.UsageReporter.Instance.ReportUIComplexity(
                _nodeCount, _depth);
        }
    }
}
