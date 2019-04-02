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

namespace VerificationRoutines
{
    [Verification(
        "ScreenReader",
        "Emulates a screen reader, except in text mode.",
        Group = "Consistency",
        Priority = 3)]
    public partial class SimpleScreenReader : UserControl, IVerificationRoutine, ISerializeTree, ICache
    {
        private List<SerializableScreenReader> _screenReaderElements = new List<SerializableScreenReader>();

        // We will only collect the strings from the TreeViewPanel so no need here, return an empty set.
        private List<string> _elementsVerifiedQueryString = new List<string>();

        GraphicsHelper _graphicsHelper;
        
        public SimpleScreenReader()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            _graphicsHelper = null;
            _screenReaderElements = null;
            _elementsVerifiedQueryString.Clear();
        }

        [XmlIgnore]
        public List<string> ElementsVerifiedQueryString
        {
            get { return _elementsVerifiedQueryString; }
            set { _elementsVerifiedQueryString = value; }
        }

        public string GetSerializedTree()
        {
            string ret = "";
            if (_screenReaderElements != null)
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
                Panel panel = graphics.CreateUIPanel("MSAA Screen reader", this.GetType());
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

        private void TraverseTree(MsaaElement element)
        {
            List<AccState> states = new List<AccState>(element.GetStates());
            List<AccRole> badRoles = new List<AccRole>(new AccRole[] { AccRole.Border, AccRole.Animation, AccRole.Graphic,
                AccRole.Grouping, AccRole.Grip, AccRole.Separator, AccRole.Sound, AccRole.Whitespace, AccRole.Client, 
                AccRole.Window });

            // if something is invisible offscreen or disabled assume that all its children are/should be as well.
            if (!states.Contains(AccState.Invisible) && !states.Contains(AccState.OffScreen) && !states.Contains(AccState.Unavailable))
            {	
                if ((!string.IsNullOrEmpty(element.GetName()) || !string.IsNullOrEmpty(element.GetValue())) && !badRoles.Contains(element.GetRole()))
                {
                    _screenReaderElements.Add(new SerializableScreenReader(element));
                }
                
                List<MsaaElement> children = element.Children;
                foreach (MsaaElement child in children)
                {
                    TraverseTree(child);
                }
            }
        }

        public void Execute(IntPtr hwnd, AccCheck.Logging.ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            if (AllowUI)
            {
                _graphicsHelper = graphics;
                
                _screenReaderElements = new List<SerializableScreenReader>();
                TestLogger.Logger = logger;
                TestLogger.RootHwnd = hwnd;
                
                MsaaElement root = MsaaElement.CacheMsaaTree(hwnd);
                TraverseTree(root);
                Panel panel = graphics.CreateUIPanel("Screen reader", this.GetType());

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
                String sValue = element.Value;
                if (String.IsNullOrEmpty(sValue))
                {
                    sValue = "";
                }
                else
                {
                    sValue = "Value: " + sValue;
                }
                
                _screenReaderOutput.Items.Add(String.Format("{0} {1} {2}\n", element.Role, element.Name, sValue));
            }
            
            //  give the list a default selection
            _screenReaderOutput.SelectedIndex = _screenReaderOutput.Items.Count - 1;

            if (_screenReaderOutput.Items.Count == 0)
            {
                _visualize.Enabled = false;
            }
            else
            {
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
                _graphicsHelper.Visualize(new VisualizeInfo(el.Name, el.Role, el.State, el.Value, el.BoundingRectangle));
            }
        }
    }


    public class SerializableScreenReader
    {
        string _name;
        string _role;
        string _state;
        string _value;
        Rectangle _boundingRectangle;

        
        public SerializableScreenReader()
        {
        }
        
        public SerializableScreenReader(MsaaElement element)
        {
            _name = element.GetName();
            _role = element.GetRole().ToString();
            _state = element.GetStatesString();
            _value = element.GetValue();
            _boundingRectangle = element.GetBoundingBox();
        }
        
       
        public string Name
        {
            get {return _name;}
            set {_name = value;}
        }
            
        public string Role
        {
            get {return _role;}
            set {_role = value;}
        }
            
        public string State
        {
            get {return _state;}
            set {_state = value;}
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
