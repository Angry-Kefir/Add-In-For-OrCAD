using System;
using Microsoft.Office.Tools.Ribbon;

namespace AddInForOrcad
{
    public partial class SampleRibbon
    {
        public event Action CreateTableClick;
        public event Action CreateEmptyTableClick;
        public event Action AddLineClick;
        public event Action DelLineClick;

        private void SampleRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void CreateTable_Click(object sender, RibbonControlEventArgs e)
        {
            if (CreateTableClick != null)
            {
                CreateTableClick();
            }
        }

        private void CreateEmptyTable_Click(object sender, RibbonControlEventArgs e)
        {
            if (CreateEmptyTableClick != null)
            {
                CreateEmptyTableClick();
            }
        }

        private void AddLine_Click(object sender, RibbonControlEventArgs e)
        {
            if (AddLineClick != null)
            {
                AddLineClick();
            }
        }

        private void DelLine_Click(object sender, RibbonControlEventArgs e)
        {
            if (DelLineClick != null)
            {
                DelLineClick();
            }
        }
    }
}
