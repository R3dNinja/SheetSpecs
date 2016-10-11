using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk;
using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;

namespace ManageSheetSpecs
{
    public partial class SheetSelector : System.Windows.Forms.Form
    {
        private UIApplication uiApp;
        private int selectedButton;
        public string ReturnValue1 { get; set; }

        public SheetSelector(UIApplication uiapp)
        {
            InitializeComponent();
            uiApp = uiapp;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            selectedButton = 1;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            selectedButton = 2;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            selectedButton = 3;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Result test = Command.thisCommand.FindSelectedTitelBlock(uiApp, selectedButton);
            if (test == Result.Failed)
            {
                this.DialogResult = DialogResult.Abort;
                Close();
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Abort;
            Close();
        }
    }
}
