using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Esquire.Common.ViewModels;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.ViewModels
{
    public class CombineSectionsViewModel : ViewModelBase
    {
        private Word.Document _document;

        public Word.Document AssociatedDocument
        {
            get
            {
                return _document;
            }
            set
            {
                //if (_document != null && value != _document)
                //{
                    _document = value;
                    OnPropertyChanged("AssociatedDocument");
                    // wrong, revise
                    Init();
                //}
            }
        }

        private string _tempIndicator;
        public string TempIndicator
        {
            get {  return _tempIndicator; }
            set
            {
                _tempIndicator = value;
                OnPropertyChanged("TempIndicator");
            }
        }

        private void Init()
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(AssociatedDocument);
            vstoDoc.SelectionChange += new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
        }

        private void ThisDocument_SelectionChange(object sender, SelectionEventArgs e)
        {
            TempIndicator = DateTime.Now.ToLongTimeString();
        }
    }
}
