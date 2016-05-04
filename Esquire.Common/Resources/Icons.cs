using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Esquire.Common.Resources
{
    public class Icons
    {
        /// <summary>
        /// Returns an icon embedded within the assembly as a bitmap the size of the default icon size.
        /// </summary>
        /// <param name="name">Name of the icon</param>
        /// <param name="assembly">Assembly the icon is embedded within</param>
        /// <returns></returns>
        public static Bitmap GetBitmap(Assembly assembly, string name)
        {
            // http://msdn.microsoft.com/en-us/library/windows/desktop/dd316921(v=vs.85).aspx
            // Don't use commonresource. CommonResource only contains one size
            int width = Win32.GetSystemMetrics(11); //SM_CXICON
            Icon ic = GetIconFromEmbeddedResource(assembly, name, new Size(width, width)); ;
            return ic != null ? ic.ToBitmap() : null;
        }

        /// <summary>
        /// Returns an icon embedded within the assembly as a bitmap.
        /// </summary>
        /// <param name="name">Name of the icon</param>
        /// <param name="assembly">Assembly the icon is embedded within</param>
        /// /// <param name="width">Desired Width of the bitmap</param>
        /// <returns></returns>
        public static Bitmap GetBitmap(Assembly assembly, string name, int width)
        {
            // http://msdn.microsoft.com/en-us/library/windows/desktop/dd316921(v=vs.85).aspx
            // Don't use commonresource. CommonResource only contains one size
            Icon ic = GetIconFromEmbeddedResource(assembly, name, new Size(width, width));
            return ic != null ? ic.ToBitmap() : null;
        }

        private static Icon GetIconFromEmbeddedResource(Assembly assembly, string name, Size size)
        {
            var rnames = assembly.GetManifestResourceNames();
            var tofind = name + ".ICO";
            foreach (string rname in rnames)
            {
                if (rname.EndsWith(tofind, StringComparison.CurrentCultureIgnoreCase))
                {
                    using (var stream = assembly.GetManifestResourceStream(rname))
                    {
                        return new Icon(stream, size);
                    }
                }
            }
            return null;
        }

        internal static class Win32
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
            public static extern int GetSystemMetrics(int nIndex);
        }
    }
}