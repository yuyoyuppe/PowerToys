// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using UIA=UIAutomationClient;

namespace UIAVerifications
{
    [Verification(
        "AmbiguousElements",
        "TBD",
        Group = "UIA Properties",
        Priority = 1)]
    public class AmbiguousElements : IVerificationRoutine
    {
        static LogMessage DuplicateSiblingNames = new LogMessage("Duplicate sibling Name+Role \"{0}\" will cause ambiguity between elements.", "DuplicateSiblingNames", 1);
        static LogMessage DuplicateSiblingIDs = new LogMessage("Duplicate Automation ID \"{0}\" will cause ambiguity between elements.", "DuplicateSiblingIDs", 1);

               
        private void TraverseFromNode(UIA.IUIAutomationElement automationElement)
        {
            HashSet<string> siblingNames = new HashSet<string>();
            HashSet<string> siblingIDs = new HashSet<string>();

            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    //Logic check for query id collision
                    //If automationID is Not null and two Automation ids are same collision will occur, error 
                    //If automationID is null and Name and Role are identical for two elements then collision will occur

                    string automationID = children.GetElement(i).CurrentAutomationId;
                    string nameRole = children.GetElement(i).CachedName;
                    nameRole = nameRole + "+" + children.GetElement(i).CachedControlType;

                    if (!(String.IsNullOrEmpty(automationID)))
                    {
                        if (siblingIDs.Contains(automationID))
                        {
                            LogHelper.Log(EventLevel.Error, DuplicateSiblingIDs, this.GetType(), automationElement, automationID);
                        }
                        else
                        {
                            siblingIDs.Add(automationID);
                        }
                    }
                    else
                    {
                        if (siblingNames.Contains(nameRole))
                        {
                            LogHelper.Log(EventLevel.Error, DuplicateSiblingNames, this.GetType(), automationElement, nameRole);
                        }
                        else
                        {
                            siblingNames.Add(nameRole);
                        }
                    }   
                    TraverseFromNode(children.GetElement(i));
                }
            }
        }
        
        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;

            TraverseFromNode(UIAGlobalContext.CacheUIATree(hwnd));
        }
    }
}
