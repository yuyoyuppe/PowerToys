// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

//
// Description: This class is used to show the bounding rectangle.
//

using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AccCheck;

namespace AccCheckUI
{
    internal class HighlightRectangle
    {
        //----------------------------------------------------------------
        //
        // Constructor
        //
        //----------------------------------------------------------------

        #region Constructor

        public HighlightRectangle()
        {
            _show = false;
            _width = 3;

            _leftForm = new Form();
            _topForm = new Form();
            _rightForm = new Form();
            _bottomForm = new Form();

            Form[] forms = { _leftForm, _topForm, _rightForm, _bottomForm };
            foreach (Form form in forms)
            {
                form.FormBorderStyle = FormBorderStyle.None;
                form.ShowInTaskbar = false;
                form.TopMost = true;
                form.Visible = false;
                form.Left = 0;
                form.Top = 0;
                form.Width = 1;
                form.Height = 1;
                form.Opacity = 0.50;

                int style = Win32API.GetWindowLong(form.Handle, Win32API.GWL_EXSTYLE);
                Win32API.SetWindowLong(form.Handle, Win32API.GWL_EXSTYLE, (int)(style | Win32API.WS_EX_TOOLWINDOW));
            }

            Color = Color.Red; // set via property to ensure rects get updated
        }

        #endregion

        //----------------------------------------------------------------
        //
        // Public Properties
        //
        //----------------------------------------------------------------

        #region Public Properties

        public bool Visible
        {
            set
            {
                if (_show != value)
                {
                    _show = value;
                    if (_show)
                    {
                        Layout();
                        Show(true);
                    }
                    else
                    {
                        Show(false);
                    }
                }
            }
        }

        public Rectangle Location
        {
            set
            {
                _location = value;
                Layout();
            }
        }

        public Color Color
        {
            set
            {
                _color = value;
                this._leftForm.BackColor = value;
                this._topForm.BackColor = value;
                this._rightForm.BackColor = value;
                this._bottomForm.BackColor = value;
            }
        }

        #endregion

        //----------------------------------------------------------------
        //
        // Private Methods
        //
        //----------------------------------------------------------------

        #region Private Methods

        private void Show(bool show)
        {
            if (show)
            {
                Win32API.ShowWindow(_leftForm.Handle, Win32API.SW_SHOWNA);
                Win32API.ShowWindow(_topForm.Handle, Win32API.SW_SHOWNA);
                Win32API.ShowWindow(_rightForm.Handle, Win32API.SW_SHOWNA);
                Win32API.ShowWindow(_bottomForm.Handle, Win32API.SW_SHOWNA);
            }
            else
            {
                this._leftForm.Hide();
                this._topForm.Hide();
                this._rightForm.Hide();
                this._bottomForm.Hide();
            }
        }

        private void Layout()
        {
            // Use SetWindowPos instead of changing the location via form properties: this allows us
            // to also specify HWND_TOPMOST. Using Form.TopMost=true to do this has the side-effect
            // of activating the windows, losing the focus.
            // Make the rect 1 pixel inside the rect so for full screen apps at the edge or the desktop 
            // you get at least a hint of where the rect is.
            Win32API.SetWindowPos(_leftForm.Handle, Win32API.HWND_TOPMOST,
                        _location.Left - _width, _location.Top, _width + 1, _location.Height, Win32API.SWP_NOACTIVATE);

            Win32API.SetWindowPos(_topForm.Handle, Win32API.HWND_TOPMOST,
                        _location.Left - _width, _location.Top - _width, _location.Width + 2 * _width, _width + 1, Win32API.SWP_NOACTIVATE);

            Win32API.SetWindowPos(_rightForm.Handle, Win32API.HWND_TOPMOST,
                        _location.Left + _location.Width  - 1, _location.Top, _width + 1, _location.Height, Win32API.SWP_NOACTIVATE);

            Win32API.SetWindowPos(_bottomForm.Handle, Win32API.HWND_TOPMOST,
                        _location.Left - _width, _location.Top + _location.Height - 1, _location.Width + 2 * _width, _width + 1, Win32API.SWP_NOACTIVATE);
        }

        #endregion

        //----------------------------------------------------------------
        //
        // Private Fields
        //
        //----------------------------------------------------------------
        #region Private Fields

        private bool _show;
        private int _width;
        private Rectangle _location;
        private Color _color;

        private Form _leftForm;
        private Form _topForm;
        private Form _rightForm;
        private Form _bottomForm;

        #endregion
    }
}
