using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Win32;


namespace CSPSolutions.COM.Keyboard
{
    /// <summary>
    ///  Define methods in the interface so that they appear, when called through com
    /// </summary>
    [Guid("18D67839-31E0-46E0-BFE1-4CE201E605FD"),
     ProgId("CSPSolutions.KeyboardManager"),
     ClassInterface(ClassInterfaceType.None)]
    public class KeyboardManager : IKeyboardManager, IDisposable
    {
        #region Types

        private delegate IntPtr HookHandlerDelegate(int nCode, IntPtr wParam, ref KBHookStruct lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBHookStruct
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public int dwExtraInfo;
        }

        #endregion

        #region Events

        #endregion

        #region Declarations

        #region Variables

        private static IntPtr _hookID = IntPtr.Zero;
        private static readonly HookHandlerDelegate _proc = KeyboardHookHandler;
        private bool _disposed = false;

        #endregion

        #region Constants

        /// <summary>
        /// Activate the layout
        /// </summary>
        private const uint KLF_ACTIVATE = 1; 
        private const int KL_NAMELENGTH = 9; // length of the keyboard buffer
     

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private const string SystemPolicyKeyPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System";
        private const string TaskManagerKeyName = "DisableTaskMgr";

        #endregion

        #endregion

        #region Constructors

        #endregion

        #region Methods

        #region Public

        /// <summary>
        /// Disable system keys
        /// </summary>
        public void DisableSystemKeys()
        {
            try
            {
                _hookID = SetHook(_proc);

                this.EnableCtrlAltDelete(false);
            }
            catch
            {
                if (_hookID != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookID);
                }
            }           
        }

        /// <summary>
        /// Enable system keys
        /// </summary>
        public void EnableSystemKeys()
        {
            try
            {
                if (_hookID != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookID);

                    _hookID = IntPtr.Zero;

                    this.EnableCtrlAltDelete(true);
                }
            }
            catch
            {
            }
        }

        public void EnableKeyboardLanguage(string languageCode)
        {
            try
            {
                if (!string.IsNullOrEmpty(languageCode))
                {
                    LoadKeyboardLayout(languageCode, KLF_ACTIVATE);
                }
            }
            catch
            {

            }
        }

        public void EnableEnglishUS()
        {
            try
            {
                LoadKeyboardLayout(LanguageCodes.EN_US, KLF_ACTIVATE);
            }
            catch
            {
               
            }
        }

        public void EnableArabicLB()
        {
            try
            {
                LoadKeyboardLayout(LanguageCodes.AR_LB, KLF_ACTIVATE);
            }
            catch
            {

            }
        }

        #endregion

        #region Private

        #region External Methods

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookHandlerDelegate lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, ref KBHookStruct lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern long LoadKeyboardLayout(string pwszKLID, uint Flags);

        [DllImport("user32.dll")]
        private static extern long GetKeyboardLayoutName(System.Text.StringBuilder pwszKLID);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        private static extern int GetKeyboardLayoutList(int size, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] hkls);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr ActivateKeyboardLayout(IntPtr hkl, int uFlags);

        #endregion

        public static ushort MAKELANGID(ushort PrimaryLanguage, ushort SubLanguage)
        {
            return Convert.ToUInt16((SubLanguage << 10) | PrimaryLanguage);
        }


        private static IntPtr SetHook(HookHandlerDelegate proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            {
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
        }

        private static IntPtr KeyboardHookHandler(int nCode, IntPtr wParam, ref KBHookStruct lParam)
        {
            Keys key = Keys.None;

            if (nCode == 0)
            {
                try
                {
                    key = (Keys)lParam.vkCode;
                }
                catch { }

                if (((lParam.vkCode == 0x09) && (lParam.flags == 0x20)) ||  // Alt+Tab
                ((lParam.vkCode == 0x1B) && (lParam.flags == 0x20)) ||      // Alt+Esc
                ((lParam.vkCode == 0x1B) && (lParam.flags == 0x00)) ||      // Ctrl+Esc

                ((lParam.vkCode == 0x5B) && (lParam.flags == 0x01)) ||      // Left Windows Key
                ((lParam.vkCode == 0x5C) && (lParam.flags == 0x01)) ||      // Right Windows Key


                (key == Keys.LWin || key == Keys.RWin) ||      //  Windows Key

                ((lParam.vkCode == 0x73) && (lParam.flags == 0x20)) ||      // Alt+F4
                ((lParam.vkCode == 0x20) && (lParam.flags == 0x20)) || // Alt+Space

                (key == Keys.None) || (key == Keys.Menu) || (key == Keys.Pause) ||
                (key == Keys.Help) || (key == Keys.Sleep) || (key == Keys.Apps)
                    || (key == Keys.PrintScreen) || (key == Keys.Print) 
                //||(key >= Keys.KanaMode && key <= Keys.HanjaMode) || (key >= Keys.IMEConvert && key <= Keys.IMEModeChange) ||
                //(key >= Keys.BrowserBack && key <= Keys.BrowserHome) || (key >= Keys.MediaNextTrack && key <= Keys.OemClear)

                 )
                {
                    return new IntPtr(1);
                }

            }

            return CallNextHookEx(_hookID, nCode, wParam, ref lParam);
        }
 
        //remove key
        private void EnableCtrlAltDelete(bool enable)
        {
            string keyValueInt;

            try
            {
                keyValueInt = string.Format("{0}", (enable ? 0: 1));

                using (RegistryKey regkey = Registry.CurrentUser.CreateSubKey(SystemPolicyKeyPath))
                {
                    regkey.SetValue(TaskManagerKeyName, keyValueInt);

                    regkey.Close();
                }
            }
            catch
            {

            }
        }
 

        #endregion

        #region Protected

        #endregion

        #region Protected Internal

        #endregion

        #region Internal

        #endregion

        #endregion

        #region Properties

        #endregion

        #region Memory Clean Up

        public void Dispose()
        {
            if (!this._disposed)
            {
                this._disposed = true;

                this.EnableSystemKeys();
            }
        }

        ~KeyboardManager()
        {
            this.Dispose();
        }

        #endregion
    }
}
