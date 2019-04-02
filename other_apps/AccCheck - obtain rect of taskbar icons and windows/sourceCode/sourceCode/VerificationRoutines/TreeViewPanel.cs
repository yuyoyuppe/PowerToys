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
using Accessibility;
using AccCheck;
using AccCheck.Verification;
using AccCheck.Logging;

namespace VerificationRoutines
{
    [Verification(
        "TreeViewer",
        "Brings up a tab that lets you examine the tree.",
        NeedsUI = true,
        Group = "Manual verification",
        CanVisualize = true)]
    public partial class TreeViewPanel : UserControl, IVerificationRoutine, ISerializeTree, ICache
    {
        private GraphicsHelper _graphicsHelper;
        private SerializableMsaaTree _serializableMsaaTree;
        private List<string> _elementsVerifiedQueryString = new List<string>();

        public TreeViewPanel()
        {
            InitializeComponent();
        }

        [XmlIgnore]
        public List<string> ElementsVerifiedQueryString
        {
            get { return _elementsVerifiedQueryString; }
            set { _elementsVerifiedQueryString = value; }
        }

        public void Clear()
        {
            _graphicsHelper = null;
            _serializableMsaaTree = null;
            _elementsVerifiedQueryString.Clear();
        }
        
        public string GetSerializedTree()
        {
            string ret = "";
            if (_serializableMsaaTree != null)
            {
                Stream mStream = new MemoryStream();
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(SerializableMsaaTree));
                XmlTextWriter xmlTextWriter = new XmlTextWriter(mStream, Encoding.UTF8);
                StreamReader reader = new StreamReader(mStream);

                xmlSerializer.Serialize(xmlTextWriter, _serializableMsaaTree);
                mStream.Position = 0;

                ret = reader.ReadToEnd();
            }

            return ret;
        }

        public void SetSerializedTree(string tree, GraphicsHelper graphics)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(SerializableMsaaTree));
            using (StringReader reader = new StringReader(tree))
            {
                try
                {
                    _serializableMsaaTree = (SerializableMsaaTree)serializer.Deserialize(reader);
                }
                catch (Exception )
                {
                    // if we can't deserialize just build the tree on the current data
                }
                
                reader.Close();

                _graphicsHelper = graphics;
                Panel panel = graphics.CreateUIPanel("MSAA Tree", this.GetType());
                if (panel.InvokeRequired)
                {
                    panel.BeginInvoke(new BuildTreeDelegate(BuildTree), panel);
                }
                else
                {
                    BuildTree(panel);
                }

                VisualizationCallback vd = new VisualizationCallback(Visualize);
                _graphicsHelper.RegisterVisualizer(vd, typeof(TreeViewPanel));

            }
        }

        private void tree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode node = e.Node;
            SerializableMsaaTree msaaElement = node.Tag as SerializableMsaaTree;

            lblName.Text = msaaElement.Name;
            lblRole.Text = msaaElement.Role;
            lblValue.Text = msaaElement.Value;
            lblLocation.Text = msaaElement.BoundingRectangle.ToString();

            pbScreenshot.Image = _graphicsHelper.Screenshot;

            if (pbScreenshot.Image != null)
            {
                _graphicsHelper.HighLightRectangle(pbScreenshot.Image, msaaElement.BoundingRectangle);
            }

            pbScreenshot.Refresh();
        }

        private void TraverseFromNode(SerializableMsaaTree msaaElement, TreeNode parent)
        {
            String name = msaaElement.Name;

            if (VerificationManager.IncludePassResultsInLog)
            {
                _elementsVerifiedQueryString.Add(msaaElement.QueryString);
            }

            // it can be empty
            name = (name == String.Empty) ? "(nothing)" : name;
            TreeNode node = new TreeNode(name);

            node.Tag = msaaElement;

            List<SerializableMsaaTree> children = msaaElement.Children;
            if (children != null)
            {
                foreach (SerializableMsaaTree child in children)
                {
                    TraverseFromNode(child, node);
                }
            }
            parent.Nodes.Add(node);
        }

        delegate void BuildTreeDelegate(Panel panel);
        public void BuildTree(Panel panel)
        {
            TreeNode node = new TreeNode("root");

            tree.Nodes.Clear();

            if (_serializableMsaaTree != null)
            {
                node.Tag = _serializableMsaaTree;
                _elementsVerifiedQueryString = new List<string>();
                TraverseFromNode(_serializableMsaaTree, node);
                tree.Nodes.Add(node);
            }

            Dock = DockStyle.Fill;
            panel.Controls.Add(this);
            pbScreenshot.Image = (Image)_graphicsHelper.Screenshot;
        }

        public bool Visualize(VisualizeInfo visualizeInfo)
        {
            bool found = false;
            if (visualizeInfo != null)
            {
                SerializableMsaaTree msaaElement = LookForElement((SerializableMsaaTree)tree.Nodes[0].Tag, visualizeInfo);
                if (msaaElement != null)
                {
                    TreeNode current = tree.Nodes[0];
                    foreach (int i in msaaElement.TreePath)
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

        SerializableMsaaTree LookForElement(SerializableMsaaTree element, VisualizeInfo visualizeInfo)
        {
            SerializableMsaaTree el = null;

            if (element.Name == visualizeInfo.Name &&
                element.Role == visualizeInfo.Role &&
                element.State == visualizeInfo.State &&
                element.Value == visualizeInfo.Value &&
                element.BoundingRectangle == visualizeInfo.Location)
            {
                return element;
            }

            foreach (SerializableMsaaTree child in element.Children)
            {
                el = LookForElement(child, visualizeInfo);
                if (el != null)
                {
                    break;
                }
            }

            return el;
        }

        #region IVerificationRoutine Members

        public void Execute(IntPtr hwnd, AccCheck.Logging.ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            // The root element will be null if no verifications from this dll have been run yet
            // if that is the case there is no need to build the tree.
            if (MsaaElement.RootElement != null)
            {
                // We want to build the serialized tree even if we are running in automation
                // so the trees can be written out to the log file
                _serializableMsaaTree = new SerializableMsaaTree(MsaaElement.RootElement);

                if (AllowUI)
                {
                    TestLogger.Logger = logger;
                    TestLogger.RootHwnd = hwnd;

                    _graphicsHelper = graphics;
                    Panel panel = graphics.CreateUIPanel("Tree", this.GetType());

                    if (panel.InvokeRequired)
                    {
                        panel.BeginInvoke(new BuildTreeDelegate(BuildTree), panel);
                    }
                    else
                    {
                        BuildTree(panel);
                    }

                    if (_graphicsHelper != null)
                    {
                        VisualizationCallback vd = new VisualizationCallback(Visualize);
                        _graphicsHelper.RegisterVisualizer(vd, typeof(TreeViewPanel));
                    }

                }

            }

        }

        #endregion
    }



    public class SerializableMsaaTree
    {
        string _name;
        string _role;
        string _state;
        string _value;
        Rectangle _boundingRectangle;
        List<SerializableMsaaTree> _children;
        string _queryString = string.Empty;

        public int[] _treePath;

        public SerializableMsaaTree()
        {
        }

        public SerializableMsaaTree(MsaaElement msaaElement)
        {
            if (VerificationManager.IncludePassResultsInLog)
            {
                IQueryString queryString = VerificationManager.GetQueryStringObject();
                if (queryString != null)
                {
                    this._queryString = queryString.Generate(msaaElement.Accessible.IAccessible, msaaElement.Accessible.ChildId);
                }
            }

            _name = msaaElement.GetName();
            _role = msaaElement.GetRole().ToString();
            AccState[] states = msaaElement.GetStates();
            StringBuilder sb = new StringBuilder();
            foreach (AccState a in states)
            {
                sb.Append(Enum.GetName(typeof(AccState), a));
                sb.Append(" ");
            }
            _state = sb.ToString();
            _value = msaaElement.GetValue();
            _boundingRectangle = msaaElement.GetBoundingBox();
            _children = null;
            _treePath = msaaElement.TreePath;

            List<MsaaElement> children = msaaElement.Children;
            if (children != null)
            {
                _children = new List<SerializableMsaaTree>();
                foreach (MsaaElement element in children)
                {
                    _children.Add(new SerializableMsaaTree(element));
                }
            }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string Role
        {
            get { return _role; }
            set { _role = value; }
        }

        public string State
        {
            get { return _state; }
            set { _state = value; }
        }

        public string Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public Rectangle BoundingRectangle
        {
            get { return _boundingRectangle; }
            set { _boundingRectangle = value; }
        }

        public List<SerializableMsaaTree> Children
        {
            get { return _children; }
            set { _children = value; }
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
            get
            {
                return this._queryString;
            }
            set
            {
                this._queryString = value;
            }
        }
    }
}
