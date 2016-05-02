using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;

namespace ClientCustomization.Interfaces
{
    public interface MyCustomInteface
    {
        void OnCustomRibbonButton(IRibbonControl ctrl);
        /// TODO:
        /// Add methods for 
        /// getVisible
        /// getPressed
        /// etc...
    }
}
