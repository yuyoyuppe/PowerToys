// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

// <summary>
//     Create a customed cursor.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Disable the warning to preserve original code
#pragma warning disable CA1716
namespace MouseWithoutBorders.Class
#pragma warning restore CA1716
{
    internal struct IconInfo
    {
        internal bool FIcon;
        internal int XHotspot;
        internal int YHotspot;
        internal IntPtr HbmMask;
        internal IntPtr HbmColor;
    }

    internal class CustomCursor
    {
        private CustomCursor()
        {
        }

        internal static Cursor CreateCursor(Bitmap bmp, int xHotSpot, int yHotSpot)
        {
            IconInfo tmp = default;
            tmp.XHotspot = xHotSpot;
            tmp.YHotspot = yHotSpot;
            tmp.FIcon = false;

            // GetIconInfo(bmp.GetHicon(), ref tmp);
            tmp.HbmColor = bmp.GetHbitmap();
            tmp.HbmMask = bmp.GetHbitmap();
            return new Cursor(NativeMethods.CreateIconIndirect(ref tmp));
        }

        internal static Cursor CreateDotCursor()
        {
            Bitmap bitmap = new(32, 32, PixelFormat.Format24bppRgb);
            bitmap.MakeTransparent(Color.Black);
            Graphics g = Graphics.FromImage(bitmap);
            g.DrawLine(Pens.Gray, 0, 0, 1, 1);
            Cursor c = CreateCursor(bitmap, 0, 0);
            bitmap.Dispose();
            return c;
        }

        private static int hidding;
        private static NativeMethods.CURSORINFO ci;

        internal static void ShowFakeMouseCursor(int x, int y)
        {
            if (Setting.Values.DrawMouse)
            {
                ci.cbSize = Marshal.SizeOf(typeof(NativeMethods.CURSORINFO));
                _ = NativeMethods.GetCursorInfo(out ci);

                // The cursor is hidden or suppressed.q
                if (ci.flags != 1)
                {
                    Common.DoSomethingInTheInputCallbackThread(
                        () =>
                    {
                        Common.MouseCursorForm ??= new FrmMouseCursor();

                        try
                        {
                            Common.MouseCursorForm.Text = string.Empty;

                            if (x == int.MinValue || y == int.MinValue)
                            {
                                hidding = 3;
                            }

                            if (hidding > 0)
                            {
                                hidding--;

                                Common.MouseCursorForm.Hide();
                            }
                            else
                            {
                                Common.MouseCursorForm.Left = x + 1;
                                Common.MouseCursorForm.Top = y + 1;
                                Common.MouseCursorForm.Width = Common.MouseCursorForm.Height = 32;
                                Common.MouseCursorForm.TopMost = true;
                                Common.MouseCursorForm.Show();
                            }
                        }
                        catch (NullReferenceException)
                        {
                            Common.Log($"{nameof(Common.MouseCursorForm)} has been set to null by another thread.");
                            Common.MouseCursorForm = new FrmMouseCursor();
                        }
                        catch (ObjectDisposedException)
                        {
                            Common.Log($"{nameof(Common.MouseCursorForm)} has been disposed.");
                            Common.MouseCursorForm = new FrmMouseCursor();
                        }
                    },
                        false);

                    return;
                }
            }

            Common.DoSomethingInTheInputCallbackThread(
                () =>
            {
                if (Common.MouseCursorForm != null)
                {
                    try
                    {
                        Common.MouseCursorForm.Close();
                        Common.MouseCursorForm.Dispose();
                    }
                    catch (NullReferenceException)
                    {
                        Common.Log($"{nameof(Common.MouseCursorForm)} has already been set to null by another thread!");
                    }
                    catch (ObjectDisposedException)
                    {
                        Common.Log($"{nameof(Common.MouseCursorForm)} has already been disposed!");
                    }

                    Common.MouseCursorForm = null;
                }
            },
                false);
        }
    }
}
