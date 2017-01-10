using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn
{
    public partial class CustomDesigner
    {
        private void CustomDesigner_Load(object sender, RibbonUIEventArgs e)
        {
            this.tab1.KeyTip = "z";
            //this.button1.KeyTip = "cks1";
        }
    }
}
