using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace CSPSolutions.COM.Keyboard
{
    /// <summary>
    /// Define methods in the interface so that they appear, when called through com
    /// </summary>
    [Guid("A47D46D3-29FD-4CF7-97E9-D6D47D10C56B")]
    public interface IKeyboardManager
    {
        void DisableSystemKeys();

        void EnableSystemKeys();

        void EnableKeyboardLanguage(string languageCode);

        void EnableEnglishUS();

        void EnableArabicLB();
    }
}
