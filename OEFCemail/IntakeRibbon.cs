using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OEFCemail
{
    public partial class IntakeRibbon
    {

        private void IntakeRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButtonIntakeDisplay_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
