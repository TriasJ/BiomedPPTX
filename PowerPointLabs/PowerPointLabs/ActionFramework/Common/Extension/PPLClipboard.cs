using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// A class that provides clipboard resources for a process to use.
    /// </summary>
    public class PPLClipboard
    {
        public static PPLClipboard Instance { get; private set; }
        public bool IsLocked { get; private set; }
        private IDataObject lDataObject;
        private IntPtr _parentWindow;
        public bool AutoDismiss;

        private PPLClipboard(IntPtr parentWindow, bool autoDismiss)
        {
            IsLocked = false;
            _parentWindow = parentWindow;
            AutoDismiss = autoDismiss;
        }

        public bool IsEmpty()
        {
            return LockIfNeeded(() =>
            {
                IDataObject clipboardData = Clipboard.GetDataObject();
                return clipboardData == null || clipboardData.GetFormats().Length == 0;
            });
        }

        public System.Drawing.Image GetImage()
        {
            return LockIfNeeded(() =>
            {
                return System.Windows.Forms.Clipboard.GetImage();
            });
        }

        public StringCollection GetFileDropList()
        {
            return LockIfNeeded(() =>
            {
                return Clipboard.GetFileDropList();
            });
        }

        public List<object> LoadClipboardObjects()
        {
            return LockIfNeeded(() =>
            {
                List<object> result = new List<object>();
                if (Clipboard.ContainsImage())
                {
                    result.Add(Clipboard.GetImage());
                }
                if (Clipboard.ContainsFileDropList())
                {
                    result.Add(Clipboard.GetFileDropList());
                }
                if (Clipboard.ContainsText())
                {
                    result.Add(Clipboard.GetText());
                }
                return result;
            });
        }

        public void LockAndRelease(Action action)
        {
            LockAndRelease<object>(() =>
            {
                action();
                return null;
            });
        }

        public TResult LockAndRelease<TResult>(Func<TResult> action)
        {
            LockClipboard();
            try
            {
                return action();
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("PPLClipboard.LockAndRelease error: " + e.Message);
                return default(TResult);
            }
            finally
            {
                ReleaseClipboard();
            }
        }

        public void LockClipboard()
        {
            if (_parentWindow == IntPtr.Zero)
            {
                try
                {
                    _parentWindow = new IntPtr(Globals.ThisAddIn.Application.HWND);
                }
                catch (Exception)
                {
                }
            }

            if (IsLocked)
            {
                return;
            }

            try
            {
                int attempts = 0;
                while (!IsClipboardFree() && !OpenClipboard(_parentWindow) && attempts < 10)
                {
                    System.Threading.Thread.Sleep(50);
                    attempts++;
                }

                IsLocked = true;
            }
            catch (Exception)
            {
                IsLocked = false;
            }
        }

        public void ReleaseClipboard()
        {
            if (!IsLocked)
            {
                return;
            }

            try
            {
                CloseClipboard();
            }
            catch (Exception)
            {
            }

            IsLocked = false;
        }

        public static void Init(IntPtr parentWindow, bool autoDismiss = false)
        {
            if (Instance == null)
            {
                Instance = new PPLClipboard(parentWindow, autoDismiss);
            }
        }

        public void Teardown()
        {
            ReleaseClipboard();
            Instance = null;
        }

        // not working stably
        public void SaveClipboard()
        {
            lDataObject = Clipboard.GetDataObject();
        }

        // referred to https://stackoverflow.com/questions/6262454/c-sharp-backing-up-and-restoring-clipboard
        // not working stably
        public void RestoreClipboard()
        {
            if (lDataObject != null)
            {
                Clipboard.SetDataObject(lDataObject);
            }
            Clipboard.Flush();
            lDataObject = null;
        }

        private void LockIfNeeded(Action action)
        {
            LockIfNeeded<object>(() =>
            {
                action();
                return null;
            });
        }

        private TResult LockIfNeeded<TResult>(Func<TResult> action)
        {
            if (IsLocked)
            {
                return action();
            }
            return LockAndRelease(action);
        }

        private bool IsClipboardFree()
        {
            IntPtr hwnd = GetOpenClipboardWindow();
            return hwnd == IntPtr.Zero || hwnd == _parentWindow;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetOpenClipboardWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool CloseClipboard();

    }
}
