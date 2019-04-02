using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Windows.Forms;
using AccCheck;
using AccCheck.Verification;
using AccCheck.Logging;
using UIA = UIAutomationClient;

namespace UIAVerifications
{
    [Verification(
        "UiaScreenReader",
        "Emulates a screen reader getting data through UIA, except in text mode.",
        Group = "Consistency",
        Priority = 3)]
    public partial class SimpleScreenReader : UserControl, IVerificationRoutine, ISerializeTree, ICache
    {
        private List<SerializableScreenReader> _screenReaderElements = new List<SerializableScreenReader>();
        private List<string> _elementsVerifiedQueryString = new List<string>();

        GraphicsHelper _graphicsHelper;
        
        public SimpleScreenReader()
        {
            InitializeComponent();
        }
        
        public void Clear()
        {
            _graphicsHelper = null;
            _screenReaderElements = new List<SerializableScreenReader>();
            _elementsVerifiedQueryString.Clear();
        }

        [XmlIgnore]
        public List<string> ElementsVerifiedQueryString
        {
            get { return _elementsVerifiedQueryString; }
            set { _elementsVerifiedQueryString = value; }
        }
        
        static List<int> ignoreControlTypes = new List<int>( new int[] {
                                              UIA.UIA_ControlTypeIds.UIA_ImageControlTypeId,
                                              UIA.UIA_ControlTypeIds.UIA_ThumbControlTypeId,
                                              UIA.UIA_ControlTypeIds.UIA_WindowControlTypeId,
                                              UIA.UIA_ControlTypeIds.UIA_PaneControlTypeId,
                                              UIA.UIA_ControlTypeIds.UIA_SeparatorControlTypeId,
                                        } );
        
        public string GetSerializedTree()
        {
            string ret = "";
            if (_screenReaderElements.Count != 0)
            {
                Stream mStream = new MemoryStream();
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<SerializableScreenReader>));
                XmlTextWriter xmlTextWriter = new XmlTextWriter(mStream, Encoding.UTF8);
                StreamReader reader = new StreamReader(mStream);
            
                xmlSerializer.Serialize(xmlTextWriter, _screenReaderElements);
                mStream.Position = 0;
            
                ret = reader.ReadToEnd();
            }

            return ret;
        }

        public void SetSerializedTree(string tree, GraphicsHelper graphics)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<SerializableScreenReader>));
            using (StringReader reader = new StringReader(tree))
            {
                try
                {
                    _screenReaderElements = (List<SerializableScreenReader>)serializer.Deserialize(reader);
                }
                catch (Exception)
                {
                    // if we can't deserialize just build the tree on the current data
                }

                reader.Close();

                _graphicsHelper = graphics;
                Panel panel = graphics.CreateUIPanel("UIA Screen reader", this.GetType());
                if (panel.InvokeRequired)
                {
                    panel.BeginInvoke(new UpdateTextBoxDelegate(UpdateTextBox), panel);
                }
                else
                {
                    UpdateTextBox(panel);
                }
            }
        }

        #region IVerificationRoutine Members

        private void TraverseTree(UIA.IUIAutomationElement element)
        {
            if (element.CachedIsEnabled != 0 && element.CachedIsOffscreen == 0)
            {
                if ((!string.IsNullOrEmpty(element.CachedName) || !string.IsNullOrEmpty(UIAGlobalContext.GetElementValue(element))) && !ignoreControlTypes.Contains(element.CachedControlType))
                {
                    _screenReaderElements.Add(new SerializableScreenReader(element));
                }

                UIA.IUIAutomationElementArray children = element.GetCachedChildren();
                if (children != null)
                {
                    for (int i = 0; i < children.Length; i++)
                    {
                        TraverseTree(children.GetElement(i));
                    }
                }
            }
        }

        public void Execute(IntPtr hwnd, AccCheck.Logging.ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            if (AllowUI)
            {
                _graphicsHelper = graphics;
                _screenReaderElements.Clear();
                LogHelper.Logger = logger;
                LogHelper.RootHwnd = hwnd;

                TraverseTree(UIAGlobalContext.CacheUIATree(hwnd));
                Panel panel = graphics.CreateUIPanel("UIA Screen reader", this.GetType());

                if (panel.InvokeRequired)
                {
                    panel.BeginInvoke(new UpdateTextBoxDelegate(UpdateTextBox), panel);
                }
                else
                {
                    UpdateTextBox(panel);
                }
            }
        }

        delegate void UpdateTextBoxDelegate(Panel panel);
        private void UpdateTextBox(Panel panel)
        {
            _screenReaderOutput.Items.Clear();
            foreach (SerializableScreenReader element in _screenReaderElements)
            {
                _screenReaderOutput.Items.Add(String.Format("{0}[{1}] {2} {3}{4}{5}\n", element.LocalizedControlType, element.ControlType, element.Name, element.AccessKey, element.Value, element.ToggleState));
            }
            
            if (_screenReaderOutput.Items.Count == 0)
            {
                _visualize.Enabled = false;
            }
            else
            {
                //  give the list a default selection
                _screenReaderOutput.SelectedIndex = _screenReaderOutput.Items.Count - 1;
                _visualize.Enabled = true;
            }
            this.Dock = DockStyle.Fill;
            panel.Controls.Add(this);
        }
        #endregion

        private void visualize_Click(object sender, EventArgs e)
        {
            if (_screenReaderElements.Count > 0 && 
                _screenReaderOutput.SelectedIndex >= 0 &&
                _screenReaderOutput.SelectedIndex < _screenReaderElements.Count)
            {
                SerializableScreenReader el = _screenReaderElements[_screenReaderOutput.SelectedIndex];

                string elementValue = el.Value;
                if (!String.IsNullOrEmpty(elementValue))
                {
                    elementValue = elementValue.Substring(el.Value.IndexOf(':') + 2);

                    // The DeSerialize gets rid of tabs and spaces if that is all there is, so mimick that behavior here.
                    string[] result = elementValue.Split(new char[] {' ', '\t'}, StringSplitOptions.RemoveEmptyEntries);
                    if (result.Length == 0)
                    {
                        elementValue = "";
                    }
                }
                _graphicsHelper.Visualize(new VisualizeInfo(el.Name, el.LocalizedControlType, el.State, elementValue, el.BoundingRectangle));
            }
        }
    }
    
    public class SerializableScreenReader
    {
        string _name;
        string _localizedControlType;
        string _controlType;
        string _accessKey;
        string _state;
        string _toggleState;
        string _value;
        Rectangle _boundingRectangle;

        
        public SerializableScreenReader()
        {
        }
        
        public SerializableScreenReader(UIA.IUIAutomationElement element)
        {
            _name = element.CachedName;
            _localizedControlType = element.CachedLocalizedControlType;
            _controlType = element.CachedControlType.ToString();
            
            string accessKey = element.CachedAccessKey;
            if(string.IsNullOrEmpty(accessKey))
            {
                _accessKey = "";
            }
            else
            {
                _accessKey = "(" + accessKey + ")";
            }
            
            _state = UIAGlobalContext.GetElementState(element);
            UIA.IUIAutomationTogglePattern togglePattern = element.GetCachedPattern(UIA.UIA_PatternIds.UIA_TogglePatternId) as UIA.IUIAutomationTogglePattern;
            if (togglePattern != null)
            {
                _toggleState = " " + togglePattern.CachedToggleState.ToString();
            }
            
            _value = "";
            UIA.IUIAutomationValuePattern valuePattern = element.GetCachedPattern(UIA.UIA_PatternIds.UIA_ValuePatternId) as UIA.IUIAutomationValuePattern;
            if (valuePattern != null)
            {
                _value = valuePattern.CachedValue;
                if (String.IsNullOrEmpty(_value))
                {
                    _value = "";
                }
                else
                {
                    _value = " Value: " + _value;
                }
            }
            else
            {
                UIA.IUIAutomationRangeValuePattern rvaluePattern = element.GetCachedPattern(UIA.UIA_PatternIds.UIA_RangeValuePatternId) as UIA.IUIAutomationRangeValuePattern;
                if (rvaluePattern != null)
                {
                    _value = rvaluePattern.CachedValue.ToString();
                    if (String.IsNullOrEmpty(_value))
                    {
                        _value = "";
                    }
                    else
                    {
                        _value = " RangeValue: " + _value;
                    }
                }
            }
            UIAutomationClient.tagRECT tempRect = element.CachedBoundingRectangle;
            _boundingRectangle = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
        }
        
       
        public string Name
        {
            get {return _name;}
            set {_name = value;}
        }
            
        public string LocalizedControlType
        {
            get {return _localizedControlType;}
            set {_localizedControlType = value;}
        }
            
        public string ControlType
        {
            get {return _controlType;}
            set {_controlType = value;}
        }
            
        public string AccessKey
        {
            get {return _accessKey;}
            set {_accessKey = value;}
        }
            
        public string State
        {
            get {return _state;}
            set {_state = value;}
        }
                
        public string ToggleState
        {
            get {return _toggleState;}
            set {_toggleState = value;}
        }
                
        public string Value
        {
            get {return _value;}
            set {_value = value;}
        }
            
        public Rectangle BoundingRectangle
        {
            get {return _boundingRectangle;}
            set {_boundingRectangle = value;}
        }
            
    }
}
