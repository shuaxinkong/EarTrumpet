using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace EarTrumpet.Interop.Helpers
{
    class ShellNotifyIcon : IDisposable
    {
        public event MouseEventHandler MouseClick;
        public event EventHandler<int> MouseWheel;

        public Icon Icon
        {
            get => _icon;
            set
            {
                if (value != _icon)
                {
                    _icon = value;
                    Update();
                }
            }
        }

        public string Text
        {
            get => _text;
            set
            {
                if (value != _text)
                {
                    _text = value;
                    Update();
                }
            }
        }

        public bool Visible
        {
            get => _isVisible;
            set
            {
                if (value != _isVisible)
                {
                    _isVisible = value;
                    Update();
                }
            }
        }

        private const int WM_CALLBACKMOUSEMSG = User32.WM_USER + 1024;
        private readonly int WM_TASKBARCREATED = User32.RegisterWindowMessage("TaskbarCreated");

        private readonly Func<Guid> _getIdentity;
        private readonly Action _resetIdentity;
        private Win32Window _window;
        private bool _isCreated;
        private bool _isVisible;
        private bool _isListeningForInput;
        private Icon _icon;
        private string _text;
        private RECT _iconLocation;
        private InputHelper.MouseInputState _cursorInfo;

        public ShellNotifyIcon(Func<Guid> getIdentity, Action resetIdentity)
        {
            _getIdentity = getIdentity;
            _resetIdentity = resetIdentity;
            _window = new Win32Window();
            _window.Initialize(WndProc);
        }

        public void SetFocus()
        {
            Trace.WriteLine("ShellNotifyIcon SetFocus");
            var data = MakeData();
            if (!Shell32.Shell_NotifyIconW(Shell32.NotifyIconMessage.NIM_SETFOCUS, ref data))
            {
                Trace.WriteLine($"ShellNotifyIcon NIM_SETFOCUS Failed: {(uint)Marshal.GetLastWin32Error()}");
            }
        }

        public void Dispose()
        {
            if (_isVisible && _isCreated)
            {
                Visible = false;
            }

            _text = null;

            if (_window != null)
            {
                _window.Dispose();
                _window = null;
            }

            if (_icon != null)
            {
                _icon.Dispose();
                _icon = null;
            }
        }

        private void Update()
        {
            var data = MakeData();

            if (_isVisible)
            {
                if (_isCreated)
                {
                    if (!Shell32.Shell_NotifyIconW(Shell32.NotifyIconMessage.NIM_MODIFY, ref data))
                    {
                        Trace.WriteLine($"ShellNotifyIcon Update NIM_MODIFY Failed: {(uint)Marshal.GetLastWin32Error()}");
                    }
                }
                else
                {
                    // If the add operation fails (due identity mismatch), reset and try again.
                    if (!Shell32.Shell_NotifyIconW(Shell32.NotifyIconMessage.NIM_ADD, ref data))
                    {
                        Trace.WriteLine($"ShellNotifyIcon Update NIM_ADD Failed {(uint)Marshal.GetLastWin32Error()}");

                        _resetIdentity();
                        data = MakeData();
                        if (!Shell32.Shell_NotifyIconW(Shell32.NotifyIconMessage.NIM_ADD, ref data))
                        {
                            Trace.WriteLine($"### ShellNotifyIcon Update NIM_ADD Failed: {(uint)Marshal.GetLastWin32Error()} ###");
                        }
                    }
                    _isCreated = true;

                    data.uTimeoutOrVersion = Shell32.NOTIFYICON_VERSION_4;
                    if (!Shell32.Shell_NotifyIconW(Shell32.NotifyIconMessage.NIM_SETVERSION, ref data))
                    {
                        Trace.WriteLine($"ShellNotifyIcon Update NIM_SETVERSION Failed: {(uint)Marshal.GetLastWin32Error()}");
                    }
                }
            }
            else if (_isCreated)
            {
                if (!Shell32.Shell_NotifyIconW(Shell32.NotifyIconMessage.NIM_DELETE, ref data))
                {
                    Trace.WriteLine($"ShellNotifyIcon Update NIM_DELETE Failed: {(uint)Marshal.GetLastWin32Error()}");
                }
                _isCreated = false;
            }
        }

        private void WndProc(Message msg)
        {
            if (msg.Msg == WM_CALLBACKMOUSEMSG)
            {
                CallbackMsgWndProc(msg);
            }
            else if (msg.Msg == WM_TASKBARCREATED)
            {
                // Shell has restarted
                Update();
            }
            else if (msg.Msg == User32.WM_INPUT)
            {
                if (InputHelper.ProcessMouseInputMessage(msg.LParam, ref _cursorInfo))
                {
                    if (IsCursorWithinNotifyIconBounds() && _cursorInfo.WheelDelta != 0)
                    {
                        MouseWheel?.Invoke(this, _cursorInfo.WheelDelta);
                    }
                }
            }
        }

        private void CallbackMsgWndProc(Message msg)
        {
            const int WM_CONTEXTMENU = 0x007B;
            const int WM_MOUSEMOVE = 0x0200;
            const int WM_LBUTTONUP = 0x0202;
            const int WM_RBUTTONUP = 0x0205;
            const int WM_MBUTTONUP = 0x0208;

            switch ((int)msg.LParam)
            {
                case (int)Shell32.NotifyIconNotification.NIN_SELECT:
                case (int)Shell32.NotifyIconNotification.NIN_KEYSELECT:
                case WM_LBUTTONUP:
                    MouseClick(this, new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0));
                    break;
                case WM_MBUTTONUP:
                    MouseClick(this, new MouseEventArgs(MouseButtons.Middle, 1, 0, 0, 0));
                    break;
                case WM_CONTEXTMENU:
                case WM_RBUTTONUP:
                    MouseClick(this, new MouseEventArgs(MouseButtons.Right, 1, 0, 0, 0));
                    break;
                case WM_MOUSEMOVE:
                    OnNotifyIconMouseMove();
                    break;
            }
        }

        private void OnNotifyIconMouseMove()
        {
            var id = new NOTIFYICONIDENTIFIER
            {
                cbSize = Marshal.SizeOf(typeof(NOTIFYICONIDENTIFIER)),
                guidItem = _getIdentity(),
            };

            if (Shell32.Shell_NotifyIconGetRect(ref id, out var location) == 0)
            {
                _iconLocation = location;

                if (User32.GetCursorPos(out var pt))
                {
                    _cursorInfo.Position = pt;
                    IsCursorWithinNotifyIconBounds();
                }
                else
                {
                    Debug.Assert(false);
                }
            }
            else
            {
                _iconLocation = default(RECT);
            }
        }

        private NOTIFYICONDATAW MakeData()
        {
            return new NOTIFYICONDATAW
            {
                cbSize = Marshal.SizeOf(typeof(NOTIFYICONDATAW)),
                hWnd = _window.Handle,
                uFlags = NotifyIconFlags.NIF_MESSAGE | NotifyIconFlags.NIF_ICON | NotifyIconFlags.NIF_TIP | NotifyIconFlags.NIF_GUID,
                uCallbackMessage = WM_CALLBACKMOUSEMSG,
                hIcon = Icon.Handle,
                szTip = Text,
                guidItem = _getIdentity(),
            };
        }

        private bool IsCursorWithinNotifyIconBounds()
        {
            bool isInBounds = _iconLocation.Contains(_cursorInfo.Position);
            if (isInBounds)
            {
                if (!_isListeningForInput)
                {
                    _isListeningForInput = true;
                    Trace.WriteLine("ShellNotifyIcon StartListening");
                    InputHelper.RegisterForMouseInput(_window.Handle);
                }
            }
            else
            {
                if (_isListeningForInput)
                {
                    _isListeningForInput = false;
                    Trace.WriteLine("ShellNotifyIcon StopListening");
                    InputHelper.UnregisterForMouseInput();
                }
            }
            return isInBounds;
        }
    }
}
