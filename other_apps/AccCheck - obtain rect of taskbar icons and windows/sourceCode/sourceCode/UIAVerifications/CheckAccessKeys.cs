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
using System.Drawing;
using UIA=UIAutomationClient;

namespace UIAVerifications
{
    [Verification(
        "CheckAccessKeysUIA",
        "Checks the Access Keys and Accelerator keys in a window",
        Group = "UIA Properties",
        Priority = 1)]
    public class CheckAccessKeysUIA : IVerificationRoutine
    {
        string[] _kbShortcutsNotToInclude = new string[] { "[null]", "[empty_string]" };

        #region LogMessages

        static LogMessage UiaDuplicateAccessKey = new LogMessage("There is a sibling Automation Element ({0}) to this one that has the same access key / accelerator key ({1})", "UiaDuplicateAccessKey", 1);

        #endregion LogMessages

        enum KeyMethod
        {
            AccessKey,
            AcceleratorKey
        };

        /// -------------------------------------------------------------------
        /// <summary>Test that will determine if the ixElement has the same 
        /// keyboard shortcut key as any if its siblings.  It will recursively 
        /// do the same for it's children.</summary>
        /// -------------------------------------------------------------------
        public void FindSiblingDuplicateKeyboardShortcut(UIA.IUIAutomationElement automationElement)
        {

            Hashtable keyTable = new Hashtable();

            bool skipNext = false;

            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    UIA.IUIAutomationElement currentElement = children.GetElement(i);

                    if (!skipNext)
                    {
                        // Add the item to the list.  If it already exists, this will log an error
                        string elementAccessKey = AddKeyboardShortcut(currentElement, KeyMethod.AccessKey, ref keyTable);
                        string elementAccelKey = AddKeyboardShortcut(currentElement, KeyMethod.AcceleratorKey, ref keyTable);

                        // And if we are a text control, need to look to see if the next controls should
                        // or should not have the same shortcut.
                        if (automationElement.CachedControlType == UIA.UIA_ControlTypeIds.UIA_TextControlTypeId)
                        {
                            UIA.IUIAutomationElement nextElement = null;

                            if (i + 1 < children.Length)
                            {
                                nextElement = children.GetElement(i + 1);
                            }
                            // If the static has a nextSibling, then look to see it is should have the 
                            // same keyboard shortcut
                            if (nextElement != null)
                            {
                                // If the names are the same, then we should have the same accessKey, if 
                                // we do, then skip this next contol and don't add the accessKey as a duplicate.
                                if (currentElement.CachedName == nextElement.CachedName)
                                {
                                    string nextElementAccessKey = GetKeyboardShortcutToLower(nextElement, KeyMethod.AccessKey);
                                    string nextElementAccelKey = GetKeyboardShortcutToLower(nextElement, KeyMethod.AcceleratorKey);

                                    if (elementAccessKey == nextElementAccessKey || elementAccelKey == nextElementAccelKey)
                                    {
                                        //Static; Names are equal; AccessKey or Accelerator keys are equal, so 
                                        //jump over this control, don't add it since the label 
                                        //defines the controls shortcut
                                        skipNext = true;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        skipNext = false;
                    }

                    // Drill into any children to make sure they have their shortcut keys 
                    // defined correctly.
                    FindSiblingDuplicateKeyboardShortcut(currentElement);
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Returns the proper accKeyboardShortcut</summary>
        /// -------------------------------------------------------------------
        private string GetKeyboardShortcutToLower(UIA.IUIAutomationElement automationElement, KeyMethod km)
        {
            string shortcut = "";
            try
            {
                if (km == KeyMethod.AccessKey)
                {
                    shortcut = automationElement.CurrentAccessKey;
                }
                else if (km == KeyMethod.AcceleratorKey)
                {
                    shortcut = automationElement.CurrentAcceleratorKey;
                }
            }
            catch (Exception e)
            {
                if (VerificationHelper.IsCriticalException(e))
                {
                    throw;
                }
            }

            // Keyboard shortcuts are optional don't want to report errors here.
            if (string.IsNullOrEmpty(shortcut))
            {
                return "";
            }
            else
            {
                return shortcut.Replace(" ", "").ToLower();
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Add the access key to the list. This will filter out known 
        /// keys that the framework uses such as [null] from the list.  If the 
        /// keyboard shortcut already exists in the list, then an error will be 
        /// reported and the string will not be added to thge list.</summary>
        /// -------------------------------------------------------------------
        string AddKeyboardShortcut(UIA.IUIAutomationElement automationElement, KeyMethod km, ref Hashtable keyTable)
        {
            string shortcutKey = GetKeyboardShortcutToLower(automationElement, km);

            if (!String.IsNullOrEmpty(shortcutKey))
            {
                if (keyTable.ContainsKey(shortcutKey))
                {
                    // We can have mutiple elements that do the same thing.  IE. Menu item can also have a popup 
                    // menu with the same command.  If they have the same name, assume that they are the same thing
                    // and in this case, is not an error.
                    string dupName = ((UIA.IUIAutomationElement)keyTable[shortcutKey]).CachedName;
                    if (dupName != automationElement.CachedName)
                    {
                        LogHelper.Log(EventLevel.Error, UiaDuplicateAccessKey, this.GetType(), automationElement, dupName, shortcutKey);
                    }
                }
                else
                {
                    keyTable.Add(shortcutKey, automationElement);
                }
            }
            return shortcutKey;
        }

        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;

            FindSiblingDuplicateKeyboardShortcut(UIAGlobalContext.CacheUIATree(hwnd));
        }
    }
}
