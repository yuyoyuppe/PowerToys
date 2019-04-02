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
     * than 50 levels deep, and gives a warning if it is more than
     * 20 levels deep. Otherwise it logs a message giving the depth
     * of the tree.
     */

    [Verification(
        "SampleCheckTreeDepth",
        "Checks that the accessibility tree isn't excessively deep.",
        Group = "Sample"
        )]

    public class CheckTreeDepth : IVerificationRoutine
    {
        private const int S_OK = 0;
        private const int ERROR_TREE_DEPTH = 50;
        private const int WARNING_TREE_DEPTH = 20;
        private int _depth = 0;

        private void TraverseTree(Accessible parent, int level)
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

            Accessible[] children;
            if (parent.Children(out children) == S_OK)
            {
                foreach (Accessible child in children)
                {
                    TraverseTree(child, level + 1);
                }
            }
        }

        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            Accessible root;
            if (Accessible.FromWindow(hwnd, out root) == S_OK)
            {
                TraverseTree(root, 0);
            }
            else
            {
                return;
            }

            if (_depth >= ERROR_TREE_DEPTH)
            {
                logger.Log(new LogEvent(EventLevel.Error, "DepthCheck", String.Format("The tree is too deep; the tree is {0} levels deep", _depth), "", System.Drawing.Rectangle.Empty, this.GetType()));
            }
            else if (_depth >= WARNING_TREE_DEPTH)
            {
                logger.Log(new LogEvent(EventLevel.Warning, "DepthCheck", String.Format("The tree might be too deep; the tree is {0} levels deep", _depth), "", System.Drawing.Rectangle.Empty, this.GetType()));
            }
            else
            {
                logger.Log(new LogEvent(EventLevel.Information, "DepthCheck", string.Format("The tree is {0} levels deep", _depth), "", System.Drawing.Rectangle.Empty, this.GetType()));
            }
        }
    }
}