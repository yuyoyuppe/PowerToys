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
using AccCheck.Logging;
using AccCheck.Verification;
using UIA=UIAutomationClient;


namespace UIAVerifications
{

    [Verification(
        "UIATreeViewer",
        "Brings up a tab that lets you examine the UIA tree.",
        NeedsUI = true,
        Group = "Manual verification",
        CanVisualize = true)]
    public partial class UIATreeView : UserControl, IVerificationRoutine, ISerializeTree, ICache
    {
        private IntPtr _hwnd;
        private GraphicsHelper _graphics;
        private UIA.IUIAutomation _automation;
        private SerializableUIATree _serializableUIATree;
        private List<string> _elementsVerifiedQueryString = new List<string>();

        public List<string> ElementsVerifiedQueryString
        {
            get { return _elementsVerifiedQueryString; }
            set { _elementsVerifiedQueryString = value; }
        }

        public UIATreeView()
        {
            _automation = new UIA.CUIAutomation();
            InitializeComponent();
        }

        public void Clear()
        {
            _hwnd = IntPtr.Zero;
            _graphics = null;
            _serializableUIATree = null;
            _elementsVerifiedQueryString.Clear();
        }

        public string GetSerializedTree()
        {
            string ret = "";
            if (_serializableUIATree != null)
            {
                Stream mStream = new MemoryStream();
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(SerializableUIATree));
                XmlTextWriter xmlTextWriter = new XmlTextWriter(mStream, Encoding.UTF8);
                StreamReader reader = new StreamReader(mStream);

                xmlSerializer.Serialize(xmlTextWriter, _serializableUIATree);
                mStream.Position = 0;

                ret = reader.ReadToEnd();
            }

            return ret;
        }

        public void SetSerializedTree(string tree, GraphicsHelper graphics)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(SerializableUIATree));
            using (StringReader reader = new StringReader(tree))
            {
                try
                {
                    _serializableUIATree = (SerializableUIATree)serializer.Deserialize(reader);
                }
                catch (Exception )
                {
                    // if we can't deserialize just build the tree on the current data
                }

                reader.Close();

                _graphics = graphics;
                Panel panel = graphics.CreateUIPanel("UIA Tree", this.GetType());
                if (panel.InvokeRequired)
                {
                    panel.BeginInvoke(new BuildTreeDelegate(BuildTree), panel);
                }
                else
                {
                    BuildTree(panel);
                }

                VisualizationCallback vd = new VisualizationCallback(Visualize);
                _graphics.RegisterVisualizer(vd, typeof(UIATreeView));

            }
        }



        private void tree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode node = e.Node;
            SerializableUIATree automationElement = node.Tag as SerializableUIATree;
            if (automationElement != null)
            {
                lblName.Text = automationElement.Name;
                lblValue.Text = automationElement.Value;
                lblControlType.Text = automationElement.ControlType;
                lblState.Text = automationElement.State;
                Rectangle rtLocation = automationElement.BoundingRectangle;
                lblLocation.Text = rtLocation.ToString();
                
                pbScreenshot.Image = _graphics.Screenshot;

                _graphics.HighLightRectangle(pbScreenshot.Image, rtLocation);
                pbScreenshot.Refresh();
            }
        }

        private void TraverseFromNode(SerializableUIATree automationElement, TreeNode parent)
        {
            String name = automationElement.Name;

            if (VerificationManager.IncludePassResultsInLog)
            {
                _elementsVerifiedQueryString.Add(automationElement.QueryString);
            }

            // it can be empty
            name = String.IsNullOrEmpty(name) ? "(nothing)" : name;
            TreeNode node = new TreeNode(name);

            node.Tag = automationElement;

            List<SerializableUIATree> children = automationElement.Children;
            if (children != null)
            {
                foreach (SerializableUIATree child in children)
                {
                    TraverseFromNode(child, node);
                }
            }
            parent.Nodes.Add(node);
        }

        public bool Visualize(VisualizeInfo visualizeInfo)
        {
            bool found = false;
            if (visualizeInfo != null)
            {
                SerializableUIATree uiaElement = LookForElement((SerializableUIATree)tree.Nodes[0].Tag, visualizeInfo);
                if (uiaElement != null)
                {
                    TreeNode current = tree.Nodes[0];
                    foreach (int i in uiaElement.TreePath)
                    {
                        if (current.Nodes.Count < i)
                        {
                            break;
                        }
                        current = current.Nodes[i];
                    }
                    tree.SelectedNode = current;
                    tree.SelectedNode.EnsureVisible();
                    found = true;
                }
            }

            return found;
        }

        SerializableUIATree LookForElement(SerializableUIATree element, VisualizeInfo visualizeInfo)
        {
            SerializableUIATree el = null;

            if (element.Name == visualizeInfo.Name &&
                element.ControlType == visualizeInfo.Role &&
                element.Value == visualizeInfo.Value &&
                element.State == visualizeInfo.State &&            
                element.BoundingRectangle == visualizeInfo.Location)
            {
                return element;
            }

            if (element.Children != null)
            {
                foreach (SerializableUIATree child in element.Children)
                {
                    el = LookForElement(child, visualizeInfo);
                    if (el != null)
                    {
                        break;
                    }
                }
            }

            return el;
        }
        
        delegate void BuildTreeDelegate(Panel panel);
        public void BuildTree(Panel panel)
        {
            TreeNode node = new TreeNode("root");

            tree.Nodes.Clear();

            if (_serializableUIATree != null)
            {
                node.Tag = _serializableUIATree;
                _elementsVerifiedQueryString = new List<string>();
                TraverseFromNode(_serializableUIATree, node);
                tree.Nodes.Add(node);
            }

            Dock = DockStyle.Fill;
            panel.Controls.Add(this);
            pbScreenshot.Image = _graphics.Screenshot;
        }

        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            UIA.IUIAutomationElement automationElement = UIAGlobalContext.LocalRootElement;
            if (automationElement != null)
            {
                // We want to build the serialized tree even if we are running in automation
                // so the trees can be written out to the log file
                _serializableUIATree = new SerializableUIATree(automationElement);

                if (AllowUI)
                {
                    _graphics = graphics;
                    _hwnd = hwnd;
                    Panel panel = graphics.CreateUIPanel("UIA Tree", this.GetType());

                    if (panel.InvokeRequired)
                    {
                        panel.BeginInvoke(new BuildTreeDelegate(BuildTree), panel);
                    }
                    else
                    {
                        BuildTree(panel);
                    }
                }

                if (_graphics != null)
                {
                    VisualizationCallback vd = new VisualizationCallback(Visualize);
                    _graphics.RegisterVisualizer(vd, typeof(UIATreeView));
                }
            }
        }
    }

    public class SerializableUIATree
    {
        String _name;
        String _controlType;
        String _value;
        String _state;
        Rectangle _boundingRectangle;
        List<SerializableUIATree> _children;
        string _queryString = string.Empty;

        public int[] _treePath;
        
        public SerializableUIATree()
        {
        }

        public SerializableUIATree(UIA.IUIAutomationElement automationElement)
        {
            Instantiate(automationElement, new int [] {0});
        }

        private SerializableUIATree(UIA.IUIAutomationElement automationElement, int [] path)
        {
            Instantiate(automationElement, path);
        }

        private void Instantiate(UIA.IUIAutomationElement automationElement, int[] path)
        {
            if (VerificationManager.IncludePassResultsInLog)
            {
                IQueryString queryString = VerificationManager.GetQueryStringObject();
                if (queryString != null)
                {
                    this._queryString = queryString.Generate(automationElement, null);
                }
            }

            _name = automationElement.CachedName;
            _controlType = automationElement.CachedLocalizedControlType;
            _value = UIAGlobalContext.GetElementValue(automationElement);
            _state = UIAGlobalContext.GetElementState(automationElement);

            UIAutomationClient.tagRECT tempRect = automationElement.CachedBoundingRectangle;
            _boundingRectangle = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
            _children = null;
            _treePath = path.Clone() as int[];


            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null)
            {
                _children = new List<SerializableUIATree>();
                for (int i = 0; i < children.Length; i++)
                {
                    List<int> newPath = new List<int>(path);
                    newPath.Add(i);
                    _children.Add(new SerializableUIATree(children.GetElement(i), newPath.ToArray()));
                }
            }
        }
        
        public string Name
        {
            get {return _name;}
            set {_name = value;}
        }
            
        public string ControlType
        {
            get {return _controlType;}
            set {_controlType = value;}
        }

        public string Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public string State
        {
            get { return _state; }
            set { _state = value; }
        }

        public Rectangle BoundingRectangle
        {
            get {return _boundingRectangle;}
            set {_boundingRectangle = value;}
        }

        public List<SerializableUIATree> Children
        {
            get {return _children;}
            set {_children = value;}
        }

        public int[] TreePath
        {
            get
            {
                return _treePath;
            }
        }

        public string QueryString
        {
            get { return _queryString; }
            set { _queryString = value; }
        }


    }

}
