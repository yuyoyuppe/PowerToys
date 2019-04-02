// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using AccCheck.Logging;

namespace AccCheck.Verification
{
    // VerificationRoutines gets a handle to a window, a logger, and a boolean
    // that tells if the app is running graphically.
    // You need to override the Execute method, and specify a metadata attribute
    public interface IVerificationRoutine
    {
        // Initialize your UI in this method. If you are UI-less, you should do all your
        // processing in this method. If you create a UI, you can make the UI event-driven,
        // but it is recommended that all your initialization code is done in this method
        void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics);
    }

    // Verification routines can cache their trees.  This can be called if the 
    // tree needs to clear the cached tree on demand.
    public interface ICache
    {
        void Clear();
    }

    // This interface allow a verification routine that exposes a tree or other UI that needs to be re populated
    // When a log file is opened the ability to serialize and de-serialize the tree. 
    public interface ISerializeTree
    {
        string GetSerializedTree();
        void SetSerializedTree(string tree, GraphicsHelper graphics);
        List<string> ElementsVerifiedQueryString { get; set; }
    }

    // If AccCheck can find an object with this interface in a dll with-in the current directory 
    // it will be used to generate strings used for suppressions that use language independent 
    // resources so that suppressions will work on UI that is in a different language then the 
    // suppressions were generated on.
    public interface IQueryString
    {
        // This method will create a string that will uniquely identify an element called a QueryString.   
        // This will work with MSAA or UIA.  MSAA clients will pass an IAccessible object and a child and 
        // UIA client will pass a IUIAutomationElement object and child id will be ignored.   This method
        // is used in the verification dll that deal with each of these accessible frameworks.
        string Generate(Object element, object childId);

        // This method will create a string that will uniquely identify an element called a QueryString.  From  
        // an hwnd.  It uses UIA to do this.
        string Generate(IntPtr hwnd);

        // This method takes a list of QueryStrings that came from one of the two methods above and a list 
        // of modules to search and return a list of QueryStrings that do not have any strings that are localized.
        // This allows these string to be used on builds of different languages.
        string[] Globalize(List<string> ids, List<string> loadedModules, out List<string> ambiguousResources);

        // Used to compare two different kids of QueryStrings
        bool Compare(string globalizedQueryString, string localQueryString);
    }
    

}
