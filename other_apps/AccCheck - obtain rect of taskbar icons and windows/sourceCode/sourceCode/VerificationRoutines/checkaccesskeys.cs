//-----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Description: This is the set of tests that are used within the AccChecker application
//-----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using Accessibility;
using System.Drawing;
using System.Diagnostics;

namespace VerificationRoutines
{
    [Verification(
        "CheckAccessKeys",
        "Check the access keys within a window",
        Group = ClassificationsGroup.Properties,
        Priority = 1)]
    public class CheckAccAccessKeys : IVerificationRoutine
    {
        string[] _kbShortcutsNotToInclude = new string[] { "[null]", "[empty_string]" };

        #region LogMessages

        static LogMessage DuplicateAccessKey = new LogMessage("There is a sibling IAccessible ({0}) to this one that has the same access key ({1})", "DuplicateAccessKey", 1);

        #endregion LogMessages


        public void Execute(IntPtr hwnd, ILogger logger, bool allowUI, GraphicsHelper graphicsHelper)
        {
            TestLogger.Logger = logger;
            FindSiblingDuplicateKeyboardShortcut(MsaaElement.CacheMsaaTree(hwnd));
        }

        /// -------------------------------------------------------------------
        /// <summary>Test that will determine if the ixElement has the same 
        /// keyboard shortcut key as any if its siblings.  It will recursively 
        /// do the same for it's children.</summary>
        /// -------------------------------------------------------------------
        [TestMethod("Find sibling duplicate keyboard shortcuts")]
        public void FindSiblingDuplicateKeyboardShortcut(MsaaElement element)
        {
            MsaaElement nextElement = null;
            MsaaElement tempElement = element;
            element = tempElement;
            Hashtable keyTable = new Hashtable();

            string elementKey = string.Empty;
            string nextElementKey = string.Empty;

            for (; element != null; element = element.GetNextSibling())
            {
                // Add the item to the list.  If it already exists, this will log an error
                AddKeyboardShortcut(element, ref keyTable);

                // And if we are a static, need to look to see if the next controls should
                // or should not have the same shortcut.
                if (element.GetWindowClass().ToLower().Contains("static"))
                {
                    // If the static has a nextSibling, then look to see it is should have the 
                    // same keyboard shortcut
                    if (null != (nextElement = element.GetNextSibling()))
                    {
                        // If the names are the same, then we should have the same accessKey, if 
                        // we do, then skip this next contol and don't add the accessKey as a duplicate.
                        if (element.GetName() == nextElement.GetName())
                        {
                            nextElementKey = GetAccKeyboardShortcutToLower(nextElement);

                            if (elementKey == nextElementKey)
                            {
                                //Static; Names are equal; AccessKey are equal, so 
                                //jump over this control, don't add it since the label 
                                //defines the controls shortcut
                                element = nextElement;
                            }
                        }
                    }
                }

                // Drill into any children to make sure they have their shortcut keys 
                // defined correctly.
                if (element != null && element.Children.Count > 0)
                    FindSiblingDuplicateKeyboardShortcut(element.Children[0]);
            }
        }
    
        /// -------------------------------------------------------------------
        /// <summary>Returns the proper accKeyboardShortcut</summary>
        /// -------------------------------------------------------------------
        private string GetAccKeyboardShortcutToLower(MsaaElement element)
        {
            string shortcut = "";
            try
            {
                shortcut = element.Accessible.KeyboardShortcut;
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
        void AddKeyboardShortcut(MsaaElement element, ref Hashtable keyTable)
        {
            string shortcutKey = GetAccKeyboardShortcutToLower(element);

            if (!String.IsNullOrEmpty(shortcutKey))
            {
                if (keyTable.ContainsKey(shortcutKey))
                {
                    // We can have mutiple elements that do the same thing.  IE. Menu item can also have a popup 
                    // menu with the same command.  If they have the same name, assume that they are the same thing
                    // and in this case, is not an error.
                    string dupName = ((MsaaElement)keyTable[shortcutKey]).GetName();
                    if (dupName != element.GetName())
                    {
                        TestLogger.Log(EventLevel.Error, DuplicateAccessKey, this.GetType(), element, dupName, shortcutKey);
                    }
                }
                else
                {
                    keyTable.Add(shortcutKey, element);
                }
            }
        }
    }
}
