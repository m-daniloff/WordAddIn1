using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    class ResourceAccess
    {
        public static Bitmap GetBitmap(string name)
        {
            // http://msdn.microsoft.com/en-us/library/windows/desktop/dd316921(v=vs.85).aspx
            // Don't use CommonResource. CommonResource only contains one size
            int width = GetSystemMetrics(11); //SM_CXICON

            Icon ic = GetIconFromEmbeddedResource(name, new Size(width, width));
            return ic.ToBitmap();
        }

        public static Bitmap GetExternalBitmap(string name)
        {
            // http://msdn.microsoft.com/en-us/library/windows/desktop/dd316921(v=vs.85).aspx
            // Don't use CommonResource. CommonResource only contains one size
            int width = GetSystemMetrics(11); //SM_CXICON

            Icon ic = GetIconFromExternalEmbeddedResource(name, new Size(width, width));
            return ic.ToBitmap();
        }


        private static Icon GetIconFromEmbeddedResource(string name, Size size)
        {
            // Reduced logging in this assembly since it's fairly stable.
            var asm = System.Reflection.Assembly.GetAssembly(typeof(ResourceAccess));
            var rnames = asm.GetManifestResourceNames();
            var tofind = name + ".ico";
            foreach (string rname in rnames)
            {
                if (rname.EndsWith(tofind, StringComparison.CurrentCultureIgnoreCase))
                {
                    using (var stream = asm.GetManifestResourceStream(rname))
                    {
                        // ReSharper disable once AssignNullToNotNullAttribute
                        return new Icon(stream, size);
                    }
                }
            }

            throw new ArgumentException("Could not find resource: " + tofind, "name");
        }

        private static Icon GetIconFromExternalEmbeddedResource(string name, Size size)
        {
            // Reduced logging in this assembly since it's fairly stable.
            //var asm = System.Reflection.Assembly.GetAssembly(typeof(ResourceAccess));
            Assembly asm = Assembly.LoadFile(@"C:\Users\mdaniloff\Documents\Visual Studio 2015\Projects\WordAddIn1\WordAddIn1\bin\Debug\ClientCustomization.dll");
            var rnames = asm.GetManifestResourceNames();
            var tofind = name + ".ico";
            foreach (string rname in rnames)
            {
                if (rname.EndsWith(tofind, StringComparison.CurrentCultureIgnoreCase))
                {
                    using (var stream = asm.GetManifestResourceStream(rname))
                    {
                        // ReSharper disable once AssignNullToNotNullAttribute
                        return new Icon(stream, size);
                    }
                }
            }

            throw new ArgumentException("Could not find resource: " + tofind, "name");
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern int GetSystemMetrics(int nIndex);
    }
}
