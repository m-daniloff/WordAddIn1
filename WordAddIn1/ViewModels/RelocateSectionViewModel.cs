using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Esquire.Common.Commanding;
using Esquire.Common.ViewModels;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.ViewModels
{
    public class RelocateSectionViewModel : ViewModelBase
    {
        private Word.Document _document;
        private Document _vstoDocument;
        public Microsoft.Office.Interop.Word.Document AssociatedDocument
        {
            get
            {
                return _document;
            }
            set
            {
                if (value != _document)
                {
                    _document = value;
                    OnPropertyChanged("AssociatedDocument");
                    // wrong, revise
                    Init();
                }
            }
        }
        private int CurrentSectionIndex { get; set; }

        private bool _firstSectionSelected;

        public bool FirstSectionSelected
        {
            get { return _firstSectionSelected; }
            set
            {
                _firstSectionSelected = value;
                OnPropertyChanged("FirstSectionSelected");
            }
        }

        private bool _lastSectionSelected;

        public bool LastSectionSelected
        {
            get { return _lastSectionSelected; }
            set
            {
                _lastSectionSelected = value;
                OnPropertyChanged("LastSectionSelected");
            }
        }

        private bool _otherSectionSelected;
        public bool OtherSectionSelected
        {
            get { return _otherSectionSelected; }
            set
            {
                _otherSectionSelected = value;
                OnPropertyChanged("OtherSectionSelected");
            }
        }

        private void Init()
        {
            _vstoDocument = Globals.Factory.GetVstoObject(AssociatedDocument);
            _vstoDocument.SelectionChange += new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
        }

        private void ThisDocument_SelectionChange(object sender, SelectionEventArgs e)
        {
            object temp = AssociatedDocument.ActiveWindow.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber];
            CurrentSectionIndex = Convert.ToInt32(temp);

            int totalSections = AssociatedDocument.Sections.Count;
        }

        #region Commands

        private RelayCommand _applyCommand;

        public ICommand ApplyCommand
        {
            get
            {
                if (_applyCommand == null)
                {
                    _applyCommand = new RelayCommand(
                        param => RelocateSection(),
                        param => CanRelocateSections);
                }
                return _applyCommand;
            }
        }
        private void RelocateSection()
        {
            SectionHelpers.RelocateSectionToTheFront(CurrentSectionIndex, AssociatedDocument);
            _vstoDocument.SelectionChange -= new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
            Close();
        }
        private bool CanRelocateSections
        {
            get
            {
                if (null == AssociatedDocument)
                    return false;
                return AssociatedDocument.Sections.Count > 1;
            }
        }

        #endregion

        public event EventHandler CloseWindow;

        private void Close()
        {
            EventHandler handler = CloseWindow;
            if (handler != null)
            {
                var e = EventArgs.Empty;
                handler(this, e);
            }
        }
    }
}
