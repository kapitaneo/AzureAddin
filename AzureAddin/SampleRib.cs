using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace AzureAddin
{
    public partial class SampleRib
    {
        private void SampleRib_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void _upload_Click(object sender, RibbonControlEventArgs e)
        {
            AzForm _form = new AzureAddin.AzForm();
            _form.Show();
        }
    }
}
