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
using System.Drawing;
using System.Reflection;
using Accessibility;
using System.Diagnostics;

namespace VerificationRoutines
{

    /// -----------------------------------------------------------------------
    /// <summary>Base class for all verification routines.  If this is
    /// used by the verification classes, the driver will be able to used the 
    /// Execute routine to call the individual test methods</summary>
    /// -----------------------------------------------------------------------
    public abstract class MsaaTraversingVerificationBase : IVerificationRoutine
    {
        protected bool _allowUI = false;
        protected GraphicsHelper _graphicsHelper;
        
        internal MethodInfo[] _testMethods;
        internal TestMethodAttribute[] _testAttributes;

        /// -------------------------------------------------------------------
        /// <summary>Execute is called by the driver.  This version of Execute
        /// will call TestManager.ExecuteTestMethodsForeachElement to execute
        /// the individual test methods</summary>
        /// -------------------------------------------------------------------
        public virtual void Execute(IntPtr hwnd, ILogger logger, bool allowUI, GraphicsHelper graphicsHelper)
        {
            _allowUI = allowUI;
            _graphicsHelper = graphicsHelper;
            
            TestLogger.Logger = logger;
            TestLogger.RootHwnd = hwnd;

            Initialize();

            try
            {
                TraverseTree(MsaaElement.CacheMsaaTree(hwnd));
            }
            catch (Exception exception)
            {
                // This should not happen, if it does, we need to fix the test logic
                TestLogger.LogTestException(EventLevel.Error, this.GetType(), exception);
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Override the property to have the test methods called only
        /// for visible elements
        /// </summary>
        /// -------------------------------------------------------------------
        protected abstract bool IgnoreInvisibleElements
        {
            get;
        }
        
        protected virtual void Initialize()
        {
        }

        /// -------------------------------------------------------------------
        /// <summary>Call this to enumerate through all the test cases within an 
        /// MsaaElement.  This will recurse child MsaaElements</summary>
        /// -------------------------------------------------------------------
        void TraverseTree(MsaaElement element)
        {
            if (IgnoreInvisibleElements & !element.IsVisible)
            {
                return;
            }
                
            CacheTestMethodInformation();

            for (int i = 0; i < _testMethods.Length; i++)
            {
                InvokeTest(_testMethods[i], _testAttributes[i], element);
            }

            foreach (MsaaElement child in element.Children)
            {
                TraverseTree(child);
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Enumerate through the verification and determine and store
        /// the information pertaining to the test.  It will save the data in the 
        /// _testMethods and _testAttributes arrays.</summary>
        /// -------------------------------------------------------------------
        void CacheTestMethodInformation()
        {
            // Only fill the collections if they have never been filled before
            if (null == _testMethods || null == _testAttributes)
            {
                List<MethodInfo> methodList = new List<MethodInfo>();
                List<TestMethodAttribute> attribList = new List<TestMethodAttribute>();

                foreach (MethodInfo methodInfo in this.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly))
                {
                    foreach (Attribute attribute in methodInfo.GetCustomAttributes(typeof(TestMethodAttribute), false))
                    {
                        methodList.Add(methodInfo);
                        attribList.Add((TestMethodAttribute)attribute);
                    }
                }
                _testMethods = methodList.ToArray();
                _testAttributes = attribList.ToArray();
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Standard method that will manage the execution of a 
        /// verification.  This is usually called from the inheriting class.</summary>
        /// -------------------------------------------------------------------
        internal void InvokeTest(MethodInfo methodInfo, TestMethodAttribute attribute, object element)
        {
            TestMethodAttribute tmAttribute = (TestMethodAttribute)attribute;
            try
            {
                methodInfo.Invoke(this, new object[] { element });
            }
            catch (TargetInvocationException targetException)
            {
                // This should not happen, if it does, we need to fix the test logic
                TestLogger.LogTestException(EventLevel.Error, this.GetType(), targetException.InnerException);
            }
        }
    }
}
