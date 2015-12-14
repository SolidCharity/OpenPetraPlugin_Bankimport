//
// DO NOT REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
//
// @Authors:
//       timop
//
// Copyright 2004-2015 by OM International
//
// This file is part of OpenPetra.org.
//
// OpenPetra.org is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// OpenPetra.org is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with OpenPetra.org.  If not, see <http://www.gnu.org/licenses/>.
//

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using SourceGrid;
using Ict.Petra.Shared;
using System.Resources;
using System.Collections.Specialized;
using GNU.Gettext;
using Ict.Common;
using Ict.Petra.Client.App.Core;
using Ict.Petra.Client.App.Core.RemoteObjects;
using Ict.Common.Controls;
using Ict.Petra.Client.CommonForms;
using Ict.Petra.Client.MFinance.Logic;

namespace Ict.Petra.Plugins.Bankimport.Client
{
    public partial class TFrmImportNewBankStatement
    {
        private const string strCSVColumnsDefault = "unused,DateEffective,Description,Amount,Currency";

        /// <summary>
        /// constant for the namespace of the bankimport plugin
        /// </summary>
        public static string PluginNamespace = "Ict.Petra.Plugins.Bankimport";

        /// <summary>
        /// use this ledger
        /// </summary>
        public Int32 LedgerNumber
        {
            set
            {
                TFinanceControls.InitialiseAccountList(ref cmbSelectBankAccount, value, true, false, true, true);

                if (cmbSelectBankAccount.Count == 0)
                {
                    MessageBox.Show(Catalog.GetString("Please create a bank account first, before importing bank statements!"),
                        Catalog.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                cmbSelectBankAccount.SetSelectedString(
                    TUserDefaults.GetStringDefault(TUserDefaults.FINANCE_BANKIMPORT_BANKACCOUNT, ""), -1);

                StringCollection list = new StringCollection();

                string[] files = Directory.GetFiles(TAppSettingsManager.ApplicationDirectory, PluginNamespace + "*.Client.dll");

                foreach (string file in files)
                {
                    string FormatName = Path.GetFileName(file).Replace(PluginNamespace, "").Replace(".Client.dll", "");

                    if (FormatName.Length > 0)
                    {
                        list.Add(FormatName);
                    }
                }

                cmbSelectPlugin.SetDataSourceStringList(list);

                cmbSelectPlugin.SetSelectedString(
                    TUserDefaults.GetStringDefault(TUserDefaults.FINANCE_BANKIMPORT_PLUGIN, ""), -1);

                bool showCSVColumns = cmbSelectPlugin.GetSelectedString().Contains("CSV");
                txtCSVColumns.Visible = showCSVColumns;

                if (showCSVColumns)
                {
                    txtCSVColumns.Text = TUserDefaults.GetStringDefault("BANKIMPORTCSV" + cmbSelectBankAccount.GetSelectedString(), strCSVColumnsDefault);
                }

                lblCSVColumns.Visible = showCSVColumns;
            }
        }

        private void SelectAccount(object sender, EventArgs e)
        {
            SelectPlugin(sender, e);
        }

        private void SelectPlugin(object sender, EventArgs e)
        {
            txtCSVColumns.Visible = cmbSelectPlugin.GetSelectedString().Contains("CSV");
            lblCSVColumns.Visible = txtCSVColumns.Visible;
            if (txtCSVColumns.Visible)
            {
                txtCSVColumns.Text = TUserDefaults.GetStringDefault("BANKIMPORTCSV" + cmbSelectBankAccount.GetSelectedString(), strCSVColumnsDefault);
            }
        }

        /// <summary>
        /// bank account code that has been selected. only valid if BtnOK was clicked
        /// </summary>
        public string FAccountCode = string.Empty;

        private void BtnOK_Click(object sender, EventArgs e)
        {
            FAccountCode = cmbSelectBankAccount.GetSelectedString();

            if (FAccountCode.Length > 0)
            {
                TUserDefaults.SetDefault(TUserDefaults.FINANCE_BANKIMPORT_PLUGIN,
                    cmbSelectPlugin.GetSelectedString());

                TUserDefaults.SetDefault(TUserDefaults.FINANCE_BANKIMPORT_BANKACCOUNT,
                    cmbSelectBankAccount.GetSelectedString());

                if (cmbSelectPlugin.GetSelectedString().Contains("CSV"))
                {
                    TUserDefaults.SetDefault("BANKIMPORTCSV" + cmbSelectBankAccount.GetSelectedString(),
                        txtCSVColumns.Text);
                }

                TUserDefaults.SaveChangedUserDefault(TUserDefaults.FINANCE_BANKIMPORT_PLUGIN);

                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }
    }
}