using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using AccCheck.Verification;
using AccCheck;

namespace VerificationRoutines
{
// This needs more work and was not speced as M1 work.  
// Removing the attribute will take it out of the list of routines but keep it building.
//     [Verification(
//        "Event viewer",
//        "Brings up a UI that lets you examine the events.",
//        Group = "Manual verification")]
    public partial class EventViewPanel : UserControl, IVerificationRoutine
    {
        private GraphicsHelper _graphicsHelper;
        private Bitmap _screenshot;
        private List<AccEvent> _accevents = new List<AccEvent>();
        private Events _events;
        
        public EventViewPanel()
        {
            InitializeComponent();
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            if (_events != null && !lvEvents.Focused)
            {
                try
                {
                    AccEvent evnt = _events.GetNextEvent();
                    while (true) // aborted by exception
                    {
                        _accevents.Add(evnt);

                        ListViewItem lvi = new ListViewItem(TimeSpan.FromMilliseconds(evnt.dtTimestamp).ToString());
                        lvi.SubItems.Add(Enum.GetName(typeof(AccEventIdentifier), evnt.evnt));
                        lvEvents.Items.Add(lvi);
                        evnt = _events.GetNextEvent();
                    }
                }
                catch (ElementNotFoundException)
                {
                    // no more elements, ignore
                }
            }
        }

        #region IVerificationRoutine Members

        public void Execute(IntPtr hwnd, AccCheck.Logging.ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            if (AllowUI)
            {
                _events = new Events(hwnd);
                _screenshot = graphics.Screenshot;
                _graphicsHelper = graphics;
                pbScreenshot.Image = (Bitmap)_screenshot.Clone();

                MsaaElement.CacheMsaaTree(hwnd);
            
                this.Dock = DockStyle.Fill;
                Panel p = graphics.CreateUIPanel("Events", this.GetType());
                p.Controls.Add(this);
            }
        }

        #endregion

        private void lvEvents_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvEvents.SelectedIndices.Count == 1)
            {
                int ixSelected = lvEvents.SelectedIndices[0];
                if (_accevents.Count > ixSelected)
                {
                    AccEvent evnt = _accevents[ixSelected];
                    pbScreenshot.Image = (Image)_screenshot.Clone();

                    Graphics g = Graphics.FromImage(pbScreenshot.Image);

                    if (evnt.msaaElement != null)
                    {
                        // dim everything else
                        Color c = Color.FromArgb(196, 0, 0, 0);
                        SolidBrush b = new SolidBrush(c);
                        Rectangle rtFull = new Rectangle(new Point(0, 0), pbScreenshot.Image.Size);
                        g.FillRectangle(b, rtFull);

                        Rectangle rtHighlight = evnt.msaaElement.GetBoundingBox();

                        // does not make any sense to highlight something empty or too big
                        if (!rtHighlight.IsEmpty && (rtFull.Contains(rtHighlight) || rtFull == rtHighlight))
                        {
                            g.DrawImage(_screenshot, rtHighlight, rtHighlight, GraphicsUnit.Pixel);
                        }
                    }
                    pbScreenshot.Refresh();
                }
            }
        }

        private void btnVisualize_Click(object sender, EventArgs e)
        {
            if (lvEvents.SelectedIndices.Count == 1)
            {
                int ixSelected = lvEvents.SelectedIndices[0];
                MsaaElement el = _accevents[ixSelected].msaaElement;
                _graphicsHelper.Visualize(new VisualizeInfo(el.GetName(), el.GetRole().ToString(), el.GetStatesString(), el.GetValue(), el.GetBoundingBox()));
            }
        }
    }
}
