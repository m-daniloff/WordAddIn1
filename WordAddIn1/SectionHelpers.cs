using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class SectionHelpers
    {
        /// <summary>
        /// The section formatting of the second section takes precedence
        /// </summary>
        /// <param name="sectionIndex"></param>
        /// <param name="document"></param>
        /// <param name="useCurrentSectionHeader"></param>
        /// <param name="useCurrentSectionFooter"></param>
        /// 
        public static void CombineSectionsSimple(int sectionIndex,
            bool useCurrentSectionHeader,
            bool useCurrentSectionFooter,
            Word.Document document)
        {
            Word.Section targetSection = document.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == sectionIndex);

            if (null == targetSection)
                return;

            // investigating
            Word.Section nextSection = document.Sections[sectionIndex + 1];
            if (useCurrentSectionHeader)
            {
                nextSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                nextSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
                nextSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
            }

            if (useCurrentSectionFooter)
            {
                nextSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                nextSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
                nextSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
            }
            
            targetSection.Range.Select();
            Word.Selection selection = document.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdExtend;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Delete(ref unit, ref count);
        }

        /// <summary>
        /// The section formatting of the first section takes precedence
        /// </summary>
        /// <param name="sectionIndex"></param>
        /// <param name="document"></param>
        /// <param name="useCurrentSectionHeader"></param>
        /// <param name="useCurrentSectionFooter"></param>
        /// 
        public static void CombineSectionsComplex(int sectionIndex,
            bool useCurrentSectionHeader,
            bool useCurrentSectionFooter,
            Word.Document document)
        {
            if (document.Sections.Count == 2)
            {
                CombineSectionsComplexSpecialCase(sectionIndex, document);
                return;
            }
            Word.Section targetSection = document.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == sectionIndex);

            if (null == targetSection)
                return;
            Word.Section nextSection = document.Sections[sectionIndex + 1];
            if (useCurrentSectionHeader)
            {
                nextSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                nextSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
                nextSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
            }

            if (useCurrentSectionFooter)
            {
                nextSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                nextSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
                nextSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
            }
            targetSection.Range.Select();
            Word.Selection selection = document.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdExtend;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Cut();

            targetSection = document.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == sectionIndex);

            if (null == targetSection)
                return;

            targetSection.Range.Select();
            selection = document.Application.Selection;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Paste();
        }

        public static void CombineSectionsComplexSpecialCase(int sectionIndex, Word.Document document)
        {
             Word.Section targetSection = document.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == sectionIndex);

            if (null == targetSection)
                return;

            targetSection.Range.Select();
            Word.Selection selection = document.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdExtend;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Cut();

            targetSection = document.Sections.Cast<Word.Section>().FirstOrDefault(section => section.Index == sectionIndex);

            if (null == targetSection)
                return;

            targetSection.Range.Select();
            selection = document.Application.Selection;
            selection.MoveRight(ref unit, ref count, Type.Missing);
            selection.MoveLeft(ref unit, ref count, ref extend);
            selection.Paste();
        }


        /// <summary>
        /// Removes section break underneath section.
        /// </summary>
        /// <param name="sectionIndex">The index of the section that will have the 
        /// section break below it removed.</param>
        /// <param name="document"></param>
        public static void DeleteSectionBreak(int sectionIndex, Word.Document document)
        {
            var section = document.Sections[sectionIndex];
            section.Range.Select();
            var selection = document.Application.Selection;
            var missing = Type.Missing;
            object unit = Word.WdUnits.wdCharacter;
            object extend = Word.WdMovementType.wdExtend;
            selection.MoveRight(ref unit, 1, missing);
            selection.MoveLeft(ref unit, 1, extend);
            selection.Delete(ref unit, 1);
        }

        public static void DeleteSection(int sectionIndex, Word.Document document)
        {
            var section = document.Sections[sectionIndex];
            section.Range.Select();
            var selection = document.Application.Selection;
            selection.Delete();
        }

        public static void RelocateSectionToTheFront(int sectionIndex, Word.Document document)
        {
            var section = document.Sections[sectionIndex];
            section.Range.Select();
            var selection = document.Application.Selection;
            selection.Cut();
            object unit = Word.WdUnits.wdStory;
            selection.HomeKey(Unit: unit);
            selection.Paste();
        }

        public static void RelocateSectionToTheEnd(int sectionIndex, Word.Document document)
        {
            var section = document.Sections[sectionIndex];
            section.Range.Select();
            var selection = document.Application.Selection;
            selection.Cut();
            object unit = Word.WdUnits.wdStory;
            selection.EndKey(Unit: unit);
            selection.Paste();
        }

        public static void RelocateSectionToLocation(int sourceIndex, int targetIndex, Word.Document document)
        {
            
        }
    }
}
