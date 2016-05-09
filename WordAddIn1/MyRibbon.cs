using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ClientCustomization.Interfaces;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace WordAddIn1
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private MyCustomInteface _customization;

        public MyRibbon()
        {
            InstantiateCustomization();
        }

        public Bitmap OnGetImage(Office.IRibbonControl control)
        {
            // Why do this? Why not use image?
            // Image only works with bitmaps
            // By doing this and not putting our actual image in the resource file, but instead
            // embedding it, we can actually use the different sizes of icons, etc.
            // http://msdn.microsoft.com/en-us/library/windows/desktop/dd316921(v=vs.85).aspx
            // http://msdn.microsoft.com/en-us/library/ms724385.aspx
            // http://stackoverflow.com/questions/4025401/selecting-the-size-of-a-system-drawing-icon
            try
            {
                return ResourceAccess.GetBitmap(control.Id);
            }
            catch (Exception)
            {
                //_logger.Error(CommonResource.Error_Icon_Not_Found, e);
            }

            // Possible future consideration: http://www.shulerent.com/2011/08/25/undocumented-office-ribbon-callback-functionality/
            return null;
        }

        public Bitmap OnGetCustomImage(Office.IRibbonControl control)
        {
            // Why do this? Why not use image?
            // Image only works with bitmaps
            // By doing this and not putting our actual image in the resource file, but instead
            // embedding it, we can actually use the different sizes of icons, etc.
            // http://msdn.microsoft.com/en-us/library/windows/desktop/dd316921(v=vs.85).aspx
            // http://msdn.microsoft.com/en-us/library/ms724385.aspx
            // http://stackoverflow.com/questions/4025401/selecting-the-size-of-a-system-drawing-icon
            try
            {
                return ResourceAccess.GetExternalBitmap(control.Id);
            }
            catch (Exception)
            {
                //_logger.Error(CommonResource.Error_Icon_Not_Found, e);
            }

            // Possible future consideration: http://www.shulerent.com/2011/08/25/undocumented-office-ribbon-callback-functionality/
            return null;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string dummy = @"D:\Dev\Dummy.txt";
            if (File.Exists(dummy))
                return GetExternalResourceText("ClientCustomization.Ribbon.xml");
            return GetResourceText("WordAddIn1.MyRibbon.xml");
        }

        public void InstantiateCustomization()
        {
            string dummy = @"D:\Dev\Dummy.txt";
            if (File.Exists(dummy))
            {
                string path =
                    @"C:\Users\mdaniloff\Documents\Visual Studio 2015\Projects\WordAddIn1\WordAddIn1\bin\Debug\";
                string assemblyName = "ClientCustomization";
                Assembly assembly = Assembly.LoadFrom(path + System.IO.Path.DirectorySeparatorChar + assemblyName + ".dll");
                _customization = assembly.CreateInstance(assemblyName + ".Main", true, BindingFlags.Default, null, null, null, null) as MyCustomInteface;
            }
        }

       

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private static string GetExternalResourceText(string resourceName)
        {
            Assembly asm = Assembly.LoadFile(@"C:\Users\mdaniloff\Documents\Visual Studio 2015\Projects\WordAddIn1\WordAddIn1\bin\Debug\ClientCustomization.dll");
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public void OnCustomRibbonButton(Office.IRibbonControl ctrl)
        {
            //var custom = new ClientCustomization.Main();

            //custom.OnCustomRibbonButton(ctrl);
            _customization.OnCustomRibbonButton(ctrl);
        }

        public void OnRibbonButton(Office.IRibbonControl ctrl)
        {
            System.Windows.Forms.MessageBox.Show("Core Create button is clicked " + ctrl.Id);
        }

        public void MergeSecondPref(Office.IRibbonControl ctl)
        {
            Globals.ThisAddIn.CombineSections();
        }

        public void MergeFirstPref(Office.IRibbonControl ctl)
        {
            Globals.ThisAddIn.CombineSectionsEx();
        }

        public void ShowSectionTools(Office.IRibbonControl ctl)
        {
            Globals.ThisAddIn.ShowSectionsTaskPane();
        }

        public bool RibbonButtonEnabled(Office.IRibbonControl ctl)
        {
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            { 
                if (Globals.ThisAddIn.Application.ActiveDocument.Sections.Count > 1)
                    return true;
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        #endregion

        public void Invalidate()
        {
            ribbon.Invalidate();
        }
    }
}
