using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace ClientCustomization
{
    public class Main : ClientCustomization.Interfaces.MyCustomInteface

    {
        public void OnCustomRibbonButton(IRibbonControl ctrl)
        {
            MessageBox.Show(ctrl.Id + " is clicked in custom dll");
        }
    }
}
