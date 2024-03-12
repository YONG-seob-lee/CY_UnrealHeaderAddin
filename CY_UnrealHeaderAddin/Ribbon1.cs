using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Security.Policy;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace CY_UnrealHeaderAddin
{
    public partial class Ribbon1
    {
        AddinFunctionLibrary Lib = new AddinFunctionLibrary();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OnClick_AddCsvButton(object sender, RibbonControlEventArgs e)
        {
            
            string RegistFilePath = Lib.GetRegistFilePath();

            if(RegistFilePath == string.Empty)
            {
                Lib.ShowMessage(AddinFunctionLibrary.EMessageType.NoRegistError, Lib.GetRegistPath());
                return;
            }
            
            Lib.MakeCsv();
        }
        private void OnClick_AddHeaderButton(object sender, RibbonControlEventArgs e)
        {
            string RegistFilePath = Lib.GetRegistFilePath();

            if (RegistFilePath == string.Empty)
            {
                Lib.ShowMessage(AddinFunctionLibrary.EMessageType.NoRegistError, Lib.GetRegistPath());
                return;
            }

            Lib.MakeHeader();
        }
        private void OnClick_AddBothButton(object sender, RibbonControlEventArgs e)
        {
            string RegistFilePath = Lib.GetRegistFilePath();

            if (RegistFilePath == string.Empty)
            {
                Lib.ShowMessage(AddinFunctionLibrary.EMessageType.NoRegistError, Lib.GetRegistPath());
                return;
            }

            Lib.MakeCsv();
            Lib.MakeHeader();
        }
        private void OnClick_RegistPathButton(object sender, RibbonControlEventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialog = new FolderBrowserDialog();
            AddinFunctionLibrary Lib = new AddinFunctionLibrary();

            if (FolderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                Lib.ShowMessage(AddinFunctionLibrary.EMessageType.AccessRegist, FolderBrowserDialog.SelectedPath);

                String EmplacePath = Lib.MakeRegistPath(Directory.GetCurrentDirectory());
                System.IO.File.WriteAllText(EmplacePath, FolderBrowserDialog.SelectedPath);
            }
        }
    }
}
