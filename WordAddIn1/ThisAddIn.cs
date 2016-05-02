using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private MyRibbon _myRibbon;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
            Word.Document doc = Application.ActiveDocument;
            Word.Section targetSection = doc.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == 1);

            if (null == targetSection)
                return;
         
            targetSection.Range.Select();
            Word.Selection selection = doc.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdExtend;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Delete(ref unit, ref count);

           
        }

        internal void CombineSectionsEx()
        {
            Word.Document doc = Application.ActiveDocument;
            Word.Section targetSection = doc.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == 2);

            if (null == targetSection)
                return;
            
            targetSection.Range.Select();
            Word.Selection selection = doc.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdExtend;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Cut();

            targetSection = doc.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == 2);

            if (null == targetSection)
                return;

            targetSection.Range.Select();
            selection = doc.Application.Selection;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Paste();
        }
       
    }
}
