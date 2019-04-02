// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using AccCheck.Logging;

namespace AccCheck
{
    public delegate bool VisualizationCallback(VisualizeInfo visualizeInfo);
    public delegate void ActivateTabPaneCallback(Type routine);
    public delegate Panel AddTabPaneCallback(string title, Type routine);

    /*
     * An instance of this class gets passed to every verification
     * routine that runs.
     */
    public class GraphicsHelper
    {
        private Point _rootOffset;
        private Bitmap _screenshot;
        public Bitmap Screenshot
        {
            get
            {
                if (_screenshot != null)
                {
                    return (Bitmap)_screenshot.Clone();
                }
                else
                {
                    return null;
                }
            }
        }

        public Point RootOffset
        {
            get
            {
                return _rootOffset;
            }
        }

        public GraphicsHelper(Bitmap screenshot, Point rootOffset)
        {
            _visualizers.Clear();
            _screenshot = screenshot;
            _rootOffset = rootOffset;
        }

        public Panel CreateUIPanel(String title, Type routine)
        {
            return AddTabDelegate(title, routine);
        }
        
        public void HighLightRectangle(Image image, Rectangle rect)
        {
            Graphics g = Graphics.FromImage(image);
            
            // dim everything else
            Color c = Color.FromArgb(124, 0, 0, 0);
            SolidBrush b = new SolidBrush(c);
            Rectangle rtFull = new Rectangle(new Point(0, 0), _screenshot.Size);
            g.FillRectangle(b, rtFull);
            
            Rectangle rtHighlight = rect;
            rtHighlight.Offset(_rootOffset);// adjust the rect to be relative to the root window
            rtHighlight.Intersect(rtFull);  // clip the boundaries to the image
            
            // does not make any sense to highlight something empty or too big
            if ((rtHighlight != Rectangle.Empty) && (rtHighlight.Width > 0 && rtHighlight.Height > 0))
            {
                g.DrawImage(_screenshot, rtHighlight, rtHighlight, GraphicsUnit.Pixel);
            }
            
            const int highlightBorderWidth = 4;
            Color highLightColor = Color.Red;
            const int opacity = 160;
            
            // Make the border bigger so that it does not lay on the actual image.
            if (rtHighlight != rtFull)
            {
                rtHighlight.Inflate(highlightBorderWidth / 2, highlightBorderWidth / 2);
            }
            
            g.DrawRectangle(new Pen(Color.FromArgb(opacity, highLightColor.R, highLightColor.G, highLightColor.B), highlightBorderWidth), rtHighlight);
        }

        public static Bitmap GrabScreenShot(IntPtr hwnd)
        {
            IntPtr sourceDC = Win32API.GetDC((IntPtr)null);
            Bitmap bitmap = null;
            try
            {
                Win32API.RECT rc = new Win32API.RECT();
                Win32API.GetWindowRect(hwnd, ref rc);
                Rectangle rect = new Rectangle(rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top);

                // grab the screenshot
                bitmap = new Bitmap(rect.Width, rect.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    IntPtr destination = graphics.GetHdc();
                    Win32API.BitBlt(destination, 0, 0, rect.Width, rect.Height, sourceDC, rect.X, rect.Y, Win32API.SRCCOPY);
                    graphics.ReleaseHdc(destination);
                }
            }
            catch (ArgumentException )
            {
                // Ignore exceptions that happen when the Rectangle has a width or height of zero
                // Other verifications will return proper errors for these so do nothing here.
            }
            finally
            {
                Win32API.ReleaseDC(hwnd, sourceDC);
            }

            return bitmap;
        }

        #region Visualization
        private struct VisualizeCallback
        {
            public VisualizationCallback visualizeDelegate;
            public Type verificationRoutine;
        }

        private static List<VisualizeCallback> _visualizers = new List<VisualizeCallback>();
        
        // These is public so it can be set by AccCheckUI.MainForm        
        public static ActivateTabPaneCallback ActivateTabDelegate;
        public static AddTabPaneCallback AddTabDelegate;

        public void RegisterVisualizer(VisualizationCallback visualizer, Type verificationRoutine)
        {
            VisualizeCallback vc = new VisualizeCallback();
            vc.visualizeDelegate = visualizer;
            vc.verificationRoutine = verificationRoutine;
            bool visualizerAlreadyRegistered = false;
            for (int i = 0; i < _visualizers.Count; i++)
            {
                if (_visualizers[i].verificationRoutine == verificationRoutine)
                {
                    visualizerAlreadyRegistered = true;
                    _visualizers[i] = vc;
                }
            }

            if (!visualizerAlreadyRegistered)
            {
                _visualizers.Add(vc);
            }
        }

        public void Visualize(VisualizeInfo visualizeInfo)
        {
            int foundItemVisualizer = -1;
            for (int i = _visualizers.Count - 1; i >= 0; i--)
            {
                try
                {
                    if (_visualizers[i].visualizeDelegate(visualizeInfo))
                    {
                        foundItemVisualizer = i;
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error: " + error);
                }
            }

            if (foundItemVisualizer >= 0)
            {
                ActivateTabDelegate(_visualizers[foundItemVisualizer].verificationRoutine);
            }
        }
        #endregion
    }

    public class VisualizeInfo
    {
        public VisualizeInfo(string name, string role, string state, string value, Rectangle location)
        {
            _name = name;  
            _role = role;     
            _state = state;   
            _value = value;  
            _location = location;
        }

        private string _name;  
        private string _role;     
        private string _state;   
        private string _value;  
        private Rectangle _location;

        public string Name
        {
            get { return _name; }
        }
        public string Role
        {
            get { return _role; }
        }
        public string State
        {
            get { return _state; }
        }
        public string Value
        {
            get { return _value; }
        }
        public Rectangle Location
        {
            get { return _location; }
        }
        
    }
}

