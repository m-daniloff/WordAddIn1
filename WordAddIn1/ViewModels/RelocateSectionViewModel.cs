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

        public RelocateSectionViewModel()
        {
            FirstSectionEnabled = true;
            LastSectionEnabled = true;
        }
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
        public int CurrentSectionIndex { get; set; }

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

        private bool _firstSectionEnabled;
        public  bool FirstSectionEnabled
        {
            get {  return _firstSectionEnabled;}
            set
            {
                _firstSectionEnabled = value;
                OnPropertyChanged("FirstSectionEnabled");
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

        private bool _lastSectionEnabled;
        public bool LastSectionEnabled
        {
            get { return _lastSectionEnabled; }
            set
            {
                _lastSectionEnabled = value;
                OnPropertyChanged("LastSectionEnabled");
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
                if (!value)
                {
                    OtherSectionNumber = 0;
                }
            }
        }

        private int _otherSectionNumber;

        public int OtherSectionNumber
        {
            get {  return _otherSectionNumber; }
            set
            {
                _otherSectionNumber = value;
                OnPropertyChanged("OtherSectionNumber");
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

            if (CurrentSectionIndex == 1)
            {
                FirstSectionSelected = false;
                FirstSectionEnabled = false;
            }
            else
            {
                FirstSectionEnabled = true;
            }

            if (CurrentSectionIndex == totalSections)
            {
                LastSectionSelected = false;
                LastSectionEnabled = false;
            }
            else
            {
                LastSectionEnabled = true;
            }
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
            if (FirstSectionSelected)
                SectionHelpers.RelocateSectionToTheFront(CurrentSectionIndex, AssociatedDocument);
            else if (LastSectionSelected)
                SectionHelpers.RelocateSectionToTheEnd(CurrentSectionIndex, AssociatedDocument);
            else 
                SectionHelpers.RelocateSectionToLocation(CurrentSectionIndex, OtherSectionNumber, AssociatedDocument);

            _vstoDocument.SelectionChange -= new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
            Close();
        }
        private bool CanRelocateSections
        {
            get
            {
                if (null == AssociatedDocument)
                    return false;
                if  (AssociatedDocument.Sections.Count < 2)
                    return false;

                if (!FirstSectionSelected && !LastSectionSelected && !OtherSectionSelected)
                    return false;

                if (OtherSectionSelected && (OtherSectionNumber == 0))
                    return false;
                return true;
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
