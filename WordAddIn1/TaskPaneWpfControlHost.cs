﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace WordAddIn1
{
    public partial class TaskPaneWpfControlHost : UserControl
    {
        public TaskPaneWpfControlHost()
        {
            InitializeComponent();
        }

        public ElementHost WpfElementHost
        {
            get
            {
                return this.wpfElementHost;
            }
        }
    }
}
