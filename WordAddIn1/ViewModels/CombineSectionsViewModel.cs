using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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
    public class CombineSectionsViewModel : ViewModelBase
    {
        private Word.Document _document;
        private Document _vstoDocument;

        public CombineSectionsViewModel()
        {
            PreviousSectionSelected = false;
            UseCurrentSelected = true;

            UseCurrentSectionFooter = true;
            UseCurrentSectionHeader = true;
        }
        public Word.Document AssociatedDocument
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

        private bool _previousSectionSelected;
        public bool PreviousSectionSelected
        {
            get
            {
                return _previousSectionSelected;
            }

            set
            {
                _previousSectionSelected = value;
                OnPropertyChanged("PreviousSectionSelected");
            }
        }

        private bool _nextSectionSelected;
        public bool NextSectionSelected
        {
            get
            {
                return _nextSectionSelected;
            }

            set
            {
                _nextSectionSelected = value;
                OnPropertyChanged("NextSectionSelected");
            }
        }

        private bool _previousSectionButtonEnabled;

        public bool PreviousSectionEnabled
        {
            get {  return _previousSectionButtonEnabled; }
            set
            {
                _previousSectionButtonEnabled = value;
                OnPropertyChanged("PreviousSectionEnabled");
                if (!_previousSectionButtonEnabled)
                {
                    PreviousSectionSelected = false;
                    NextSectionSelected = true;
                }
            }
        }

        private bool _nextSectionButtonEnabled;

        public bool NextSectionEnabled
        {
            get { return _nextSectionButtonEnabled; }
            set
            {
                _nextSectionButtonEnabled = value;
                OnPropertyChanged("NextSectionEnabled");
                if (!_nextSectionButtonEnabled)
                {
                    PreviousSectionSelected = true;
                    NextSectionSelected = false;
                }
            }
        }

        private bool _useCurrentSelected;

        public bool UseCurrentSelected
        {
            get {  return _useCurrentSelected;}
            set
            {
                _useCurrentSelected = value;
                OnPropertyChanged("UseCurrentSelected");
            }
        }

        private bool _useCurrentSectionFooter;

        public bool UseCurrentSectionFooter
        {
            get { return _useCurrentSectionFooter; }
            set
            {
                _useCurrentSectionFooter = value;
                OnPropertyChanged("UseCurrentSectionFooter");
            }
        }

        private bool _useCurrentSectionHeader;

        public bool UseCurrentSectionHeader
        {
            get {  return _useCurrentSectionHeader;}
            set
            {
                _useCurrentSectionHeader = value;
                OnPropertyChanged("UseCurrentSectionHeader");
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
            _vstoDocument = Globals.Factory.GetVstoObject(AssociatedDocument);
            _vstoDocument.SelectionChange += new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
        }

        private void ThisDocument_SelectionChange(object sender, SelectionEventArgs e)
        {
            object temp = AssociatedDocument.ActiveWindow.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber];
            CurrentSectionIndex = Convert.ToInt32(temp);

            int totalSections = AssociatedDocument.Sections.Count;

            PreviousSectionEnabled = (CurrentSectionIndex != 1);
            NextSectionEnabled = (CurrentSectionIndex != totalSections);

            TempIndicator = CurrentSectionIndex.ToString(); //DateTime.Now.ToLongTimeString();
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
                        param => CombineSections(), 
                        param => CanCombineSections);       
                }
                return _applyCommand;
            }
        }

        private bool CanCombineSections
        {
            get
            {
                // return AssociatedDocument?.Sections.Count > 1;
                if (null == AssociatedDocument)
                    return false;

                return AssociatedDocument.Sections.Count > 1;
            }
        }
            
            

        private void CombineSections()
        {
            _vstoDocument.SelectionChange -= new Microsoft.Office.Tools.Word.SelectionEventHandler(ThisDocument_SelectionChange);
            if (PreviousSectionSelected)
            {
                if (UseCurrentSelected)
                    SectionHelpers.CombineSectionsComplex(GetCurrentSectionNumber(),
                        UseCurrentSectionHeader,
                        UseCurrentSectionFooter, 
                        AssociatedDocument);
                else
                {
                    SectionHelpers.CombineSectionsSimple(GetCurrentSectionNumber(), 
                        UseCurrentSectionHeader,
                        UseCurrentSectionFooter,
                        AssociatedDocument);
                }
            }

            else
            {
                if (UseCurrentSelected)
                    SectionHelpers.CombineSectionsComplex(CurrentSectionIndex,
                        UseCurrentSectionHeader,
                        UseCurrentSectionFooter,
                        AssociatedDocument);
                else
                {
                    SectionHelpers.CombineSectionsSimple(CurrentSectionIndex,
                        UseCurrentSectionHeader,
                        UseCurrentSectionFooter,
                        AssociatedDocument);
                }
            }
            
            Close();
        }

        private int GetCurrentSectionNumber()
        {
            if (CurrentSectionIndex > 1)
                CurrentSectionIndex = CurrentSectionIndex - 1;

            return CurrentSectionIndex;
        }

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

        #endregion

    }
}
