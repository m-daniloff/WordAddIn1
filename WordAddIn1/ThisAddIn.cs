using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using WordAddIn1.ViewModels;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private MyRibbon _myRibbon;
        private CustomTaskPane _taskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentChange += Application_DocumentChange;
        }

        private void Application_DocumentChange()
        {
            _myRibbon.Invalidate();
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }



        #endregion

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _myRibbon = new MyRibbon();
            return _myRibbon;
        }


        //Sections logic
        internal void CombineSections()
        {
            int count = Application.ActiveDocument.Sections.Count;
            Word.Section s1 = Application.ActiveDocument.Sections[1];
            Word.Section s2 = Application.ActiveDocument.Sections[2];

            SectionHelpers.CombineSectionsSimple(1, Application.ActiveDocument);
        }

        internal void CombineSectionsEx()
        {
            int count = Application.ActiveDocument.Sections.Count;
            Word.Section s1 = Application.ActiveDocument.Sections[0];
            Word.Section s2 = Application.ActiveDocument.Sections[1];

            SectionHelpers.CombineSectionsComplex(1, Application.ActiveDocument);
        }

        public void ShowSectionsTaskPane()
        {
            if (_taskPane == null)
                CreateCustomTaskpane();

            if (_taskPane == null)
                throw new Exception("couldn't create a Custom Task Pane");

            _taskPane.Visible = !_taskPane.Visible;

        }

        private void CreateCustomTaskpane()
        {
            var wpfHost = new TaskPaneWpfControlHost();
            var wpfControl = new CombineSectionsControl();
            ViewModels.CombineSectionsViewModel vm = new CombineSectionsViewModel();
            vm.AssociatedDocument = Application.ActiveDocument;
            wpfControl.DataContext = vm;
            wpfHost.WpfElementHost.HostContainer.Children.Add(wpfControl);
            _taskPane = this.CustomTaskPanes.Add(wpfHost, "Sections Task Pane");
            _taskPane.Visible = false;

           vm.CloseWindow += VmOnCloseWindow;
        }

        private void VmOnCloseWindow(object sender, EventArgs eventArgs)
        {
            this.CustomTaskPanes.Remove(_taskPane);
            _taskPane = null;
            // we might've turned the document into a single section document
            _myRibbon.Invalidate();
        }
    }
}
