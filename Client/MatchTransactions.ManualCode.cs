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
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using System.Threading;
using System.Collections.Generic;
using System.Xml;
using System.Globalization;
using GNU.Gettext;
using Ict.Common;
using Ict.Common.IO;
using Ict.Common.Data; // Implicit reference
using Ict.Common.Verification;
using Ict.Common.Remoting.Shared;
using Ict.Common.Remoting.Client;
using Ict.Common.Printing;
using Ict.Petra.Client.App.Core.RemoteObjects;
using Ict.Petra.Client.App.Core;
using Ict.Petra.Shared;
using Ict.Petra.Shared.MFinance;
using Ict.Petra.Shared.MFinance.Account.Data;
using Ict.Petra.Shared.MFinance.Gift.Data;
using Ict.Petra.Plugins.Bankimport.Data;
using Ict.Petra.Client.CommonDialogs;
using Ict.Petra.Client.MFinance.Logic;
using Ict.Petra.Client.MFinance.Gui.Gift;
using Ict.Petra.Plugins.Bankimport.RemoteObjects;

namespace Ict.Petra.Plugins.Bankimport.Client
{
    /// manual methods for the generated window
    public partial class TFrmMatchTransactions
    {
        private Int32 FLedgerNumber;
        private BankImportTDS FMainDS;
        private DataView FMatchView;

        /// <summary>
        /// pass the data
        /// </summary>
        public BankImportTDS MainDS
        {
            set
            {
                this.FMainDS = new BankImportTDS();
                this.FMainDS.Merge(value);
            }
            get
            {
                return FMainDS;
            }
        }

        /// <summary>
        /// use this ledger
        /// </summary>
        public Int32 LedgerNumber
        {
            set
            {
                FLedgerNumber = value;

                ALedgerRow Ledger =
                    ((ALedgerTable)TDataCache.TMFinance.GetCacheableFinanceTable(TCacheableFinanceTablesEnum.LedgerDetails, FLedgerNumber))[0];
                txtAmount.CurrencyCode = Ledger.BaseCurrency;

                TFinanceControls.InitialiseMotivationGroupList(ref cmbMotivationGroup, FLedgerNumber, true);
                TFinanceControls.InitialiseMotivationDetailList(ref cmbMotivationDetail, FLedgerNumber, true);
                TFinanceControls.InitialiseCostCentreList(ref cmbGLCostCentre, FLedgerNumber, true, false, true, true);
                TFinanceControls.InitialiseAccountList(ref cmbGLAccount, FLedgerNumber, true, false, true, false);

                grdGiftDetails.Columns.Clear();
                grdGiftDetails.AddTextColumn(Catalog.GetString("Motivation"), FMainDS.AEpMatch.ColumnMotivationDetailCode, 100);
                grdGiftDetails.AddTextColumn(Catalog.GetString("Cost Centre"), FMainDS.AEpMatch.ColumnCostCentreCode, 150);
                grdGiftDetails.AddTextColumn(Catalog.GetString("Cost Centre Name"), FMainDS.AEpMatch.ColumnCostCentreName, 200);
                grdGiftDetails.AddCurrencyColumn(Catalog.GetString("Amount"), FMainDS.AEpMatch.ColumnGiftTransactionAmount);
                FMatchView = FMainDS.AEpMatch.DefaultView;
                FMatchView.AllowNew = false;
                grdGiftDetails.DataSource = new DevAge.ComponentModel.BoundDataView(FMatchView);
            }
        }

        /// <summary>
        /// the matchtext of the transaction that we want to match
        /// </summary>
        public string MatchText
        {
            set
            {
                DataView findtransaction = new DataView(FMainDS.AEpTransaction);
                findtransaction.RowFilter = AEpTransactionTable.GetMatchTextDBName() + " = '" + value.ToString() + "'";
                this.CurrentlySelectedTransaction = (BankImportTDSAEpTransactionRow)findtransaction[0].Row;

                string description = string.Empty;

                if (CurrentlySelectedTransaction.AccountName.Length > 0)
                {
                    description = StringHelper.AddCSV(description, CurrentlySelectedTransaction.AccountName, " - ");
                }

                if (CurrentlySelectedTransaction.Description.Length > 0)
                {
                    description = StringHelper.AddCSV(description, CurrentlySelectedTransaction.Description, " - ");
                }

                txtTransactionDescription.Text = description;

                // load selections from the a_ep_match table for the new row
                FMatchView.RowFilter = AEpMatchTable.GetMatchTextDBName() +
                                       " = '" + CurrentlySelectedTransaction.MatchText + "'";

                AEpMatchRow match = (AEpMatchRow)FMatchView[0].Row;

                if (match.Action == MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT)
                {
                    rbtGift.Checked = true;

                    txtDonorKey.Text = StringHelper.FormatStrToPartnerKeyString(match.DonorKey.ToString());

                    grdGiftDetails.SelectRowInGrid(1);
                    // grdGiftDetails.SelectRowInGrid does not seem to update the gift details, so we call that manually
                    GiftDetailsFocusedRowChanged(null, null);
                }
                else if (match.Action == MFinanceConstants.BANK_STMT_STATUS_MATCHED_GL)
                {
                    rbtGL.Checked = true;
                    DisplayGLDetails();
                }
                else if (match.Action == MFinanceConstants.BANK_STMT_STATUS_NO_MATCHING)
                {
                    rbtIgnored.Checked = true;
                }
                else
                {
                    rbtUnmatched.Checked = true;
                }

                rbtGLWasChecked = rbtGL.Checked;
                rbtGiftWasChecked = rbtGift.Checked;
            }
        }

        /// <summary>
        /// return the selected transaction
        /// </summary>
        public BankImportTDSAEpTransactionRow SelectedTransaction
        {
            get
            {
                return CurrentlySelectedTransaction;
            }
        }

        /// <summary>
        /// return the updated matches
        /// </summary>
        public DataView UpdatedMatches
        {
            get
            {
                return FMatchView;
            }
        }

        private AMotivationDetailRow GetCurrentMotivationDetail(string AMotivationGroupCode, string AMotivationDetailCode)
        {
            return (AMotivationDetailRow)FMainDS.AMotivationDetail.Rows.Find(
                new object[] { FLedgerNumber, AMotivationGroupCode, AMotivationDetailCode });
        }

        private void FilterMotivationDetail(object sender, EventArgs e)
        {
            TFinanceControls.ChangeFilterMotivationDetailList(ref cmbMotivationDetail, cmbMotivationGroup.GetSelectedString());
        }

        private void SetRecipientCostCentreAndField()
        {
            // look for the motivation detail.
            AMotivationDetailRow motivationDetailRow = GetCurrentMotivationDetail(
                cmbMotivationGroup.GetSelectedString(),
                cmbMotivationDetail.GetSelectedString());

            if (motivationDetailRow != null)
            {
                txtGiftAccount.Text = motivationDetailRow.AccountCode;
                txtGiftCostCentre.Text = motivationDetailRow.CostCentreCode;
            }
            else
            {
                txtGiftAccount.Text = string.Empty;
                txtGiftCostCentre.Text = string.Empty;
            }

            Int64 RecipientKey = Convert.ToInt64(txtRecipientKey.Text);
            FInKeyMinistryChanging++;
            TFinanceControls.GetRecipientData(ref cmbMinistry, ref txtField, RecipientKey);
            FInKeyMinistryChanging--;

            long FieldNumber = Convert.ToInt64(txtField.Text);

            txtGiftCostCentre.Text = TRemote.MFinance.Gift.WebConnectors.IdentifyPartnerCostCentre(FLedgerNumber, FieldNumber);
        }

        private void MotivationDetailChanged(System.Object sender, EventArgs e)
        {
            SetRecipientCostCentreAndField();
        }

        Int16 FInKeyMinistryChanging = 0;
        private void KeyMinistryChanged(object sender, EventArgs e)
        {
            if (FInKeyMinistryChanging > 0) // || FPetraUtilsObject.SuppressChangeDetection
            {
                return;
            }

            FInKeyMinistryChanging++;
            try
            {
                Int64 rcp = cmbMinistry.GetSelectedInt64();

                txtRecipientKey.Text = String.Format("{0:0000000000}", rcp);
            }
            finally
            {
                FInKeyMinistryChanging--;
            }
        }

        private void NewTransactionCategory(System.Object sender, EventArgs e)
        {
            // do NOT call GetValuesFromScreen to avoid disappearing transaction from the grid
            // GetValuesFromScreen();
            CurrentlySelectedMatch = null;

            rbtGLWasChecked = rbtGL.Checked;
            rbtGiftWasChecked = rbtGift.Checked;

            pnlGiftEdit.Visible = rbtGift.Checked;
            pnlGLEdit.Visible = rbtGL.Checked;

            if (rbtGift.Checked)
            {
                // select first detail
                grdGiftDetails.SelectRowInGrid(1);
                AEpMatchRow match = GetSelectedMatch();
                DisplayGiftDetails();
                txtDonorKey.Text = StringHelper.FormatStrToPartnerKeyString(match.DonorKey.ToString());
            }

            if (rbtGL.Checked)
            {
                DisplayGLDetails();
            }

            if (rbtUnmatched.Checked && (FMatchView != null))
            {
                foreach (DataRowView rv in FMatchView)
                {
                    ((AEpMatchRow)rv.Row).Action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;

                    if (!CurrentlySelectedTransaction.IsEpMatchKeyNull()
                        && (CurrentlySelectedTransaction.EpMatchKey != ((AEpMatchRow)rv.Row).EpMatchKey))
                    {
                        ((AEpMatchRow)rv.Row).Delete();
                    }
                }
            }

            if (rbtIgnored.Checked && (FMatchView != null))
            {
                foreach (DataRowView rv in FMatchView)
                {
                    ((AEpMatchRow)rv.Row).Action = MFinanceConstants.BANK_STMT_STATUS_NO_MATCHING;

                    if (!CurrentlySelectedTransaction.IsEpMatchKeyNull()
                        && (CurrentlySelectedTransaction.EpMatchKey != ((AEpMatchRow)rv.Row).EpMatchKey))
                    {
                        ((AEpMatchRow)rv.Row).Delete();
                    }
                }
            }
        }

        private BankImportTDSAEpTransactionRow CurrentlySelectedTransaction = null;
        private BankImportTDSAEpMatchRow CurrentlySelectedMatch = null;
        private bool rbtGLWasChecked = false;
        private bool rbtGiftWasChecked = false;

        /// store current selections in the a_ep_match table
        private void GetValuesFromScreen()
        {
            if (CurrentlySelectedTransaction == null)
            {
                return;
            }

            if (rbtUnmatched.Checked)
            {
                if (FMatchView != null)
                {
                    for (int i = 0; i < FMatchView.Count; i++)
                    {
                        AEpMatchRow match = (AEpMatchRow)FMatchView[i].Row;
                        match.Action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
                    }
                }

                CurrentlySelectedTransaction.MatchAction = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
            }

            if (rbtIgnored.Checked)
            {
                if (FMatchView != null)
                {
                    for (int i = 0; i < FMatchView.Count; i++)
                    {
                        AEpMatchRow match = (AEpMatchRow)FMatchView[i].Row;
                        match.Action = MFinanceConstants.BANK_STMT_STATUS_NO_MATCHING;
                    }
                }

                CurrentlySelectedTransaction.MatchAction = MFinanceConstants.BANK_STMT_STATUS_NO_MATCHING;
            }

            if (CurrentlySelectedMatch == null)
            {
                return;
            }

            if (rbtGiftWasChecked)
            {
                for (int i = 0; i < FMatchView.Count; i++)
                {
                    AEpMatchRow match = (AEpMatchRow)FMatchView[i].Row;
                    match.DonorKey = Convert.ToInt64(txtDonorKey.Text);
                    match.Action = MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT;
                }

                CurrentlySelectedTransaction.MatchAction = MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT;

                GetGiftDetailValuesFromScreen();

                // TODO: validation> calculate the sum of the gift details and check with the bank transaction amount
            }

            if (rbtGLWasChecked)
            {
                CurrentlySelectedMatch.Action = MFinanceConstants.BANK_STMT_STATUS_MATCHED_GL;
                CurrentlySelectedTransaction.MatchAction = MFinanceConstants.BANK_STMT_STATUS_MATCHED_GL;

                GetGLValuesFromScreen();
            }
        }

        private void GiftDetailsFocusedRowChanged(System.Object sender, EventArgs e)
        {
            GetGiftDetailValuesFromScreen();
            CurrentlySelectedMatch = GetSelectedMatch();
            DisplayGiftDetails();
        }

        private BankImportTDSAEpMatchRow GetSelectedMatch()
        {
            DataRowView[] SelectedGridRow = grdGiftDetails.SelectedDataRowsAsDataRowView;

            if (SelectedGridRow.Length >= 1)
            {
                return (BankImportTDSAEpMatchRow)SelectedGridRow[0].Row;
            }

            return null;
        }

        private void DisplayGiftDetails()
        {
            CurrentlySelectedMatch = GetSelectedMatch();

            if (CurrentlySelectedMatch != null)
            {
                txtAmount.NumberValueDecimal = CurrentlySelectedMatch.GiftTransactionAmount;
                txtRecipientKey.Text = StringHelper.FormatStrToPartnerKeyString(CurrentlySelectedMatch.RecipientKey.ToString());

                if (CurrentlySelectedMatch.IsMotivationGroupCodeNull())
                {
                    cmbMotivationGroup.SelectedIndex = -1;
                }
                else
                {
                    cmbMotivationGroup.SetSelectedString(CurrentlySelectedMatch.MotivationGroupCode);
                }

                if (CurrentlySelectedMatch.IsMotivationDetailCodeNull())
                {
                    cmbMotivationDetail.SelectedIndex = -1;
                }
                else
                {
                    cmbMotivationDetail.SetSelectedString(CurrentlySelectedMatch.MotivationDetailCode);
                }

                SetRecipientCostCentreAndField();
            }
            else
            {
                txtAmount.NumberValueDecimal = CurrentlySelectedTransaction.TransactionAmount;
            }

            pnlEditGiftDetail.Enabled = true;
        }

        private void GetGiftDetailValuesFromScreen()
        {
            if (CurrentlySelectedMatch != null)
            {
                CurrentlySelectedMatch.MotivationGroupCode = cmbMotivationGroup.GetSelectedString();
                CurrentlySelectedMatch.MotivationDetailCode = cmbMotivationDetail.GetSelectedString();
                CurrentlySelectedMatch.CostCentreCode = txtGiftCostCentre.Text;
                CurrentlySelectedMatch.GiftTransactionAmount = txtAmount.NumberValueDecimal.Value;
                CurrentlySelectedMatch.DonorKey = Convert.ToInt64(txtDonorKey.Text);
                CurrentlySelectedMatch.RecipientKey = Convert.ToInt64(txtRecipientKey.Text);

                FMainDS.ACostCentre.DefaultView.RowFilter = String.Format("{0}='{1}'",
                    ACostCentreTable.GetCostCentreCodeDBName(), CurrentlySelectedMatch.CostCentreCode);

                if (FMainDS.ACostCentre.DefaultView.Count == 1)
                {
                    CurrentlySelectedMatch.CostCentreName = ((ACostCentreRow)FMainDS.ACostCentre.DefaultView[0].Row).CostCentreName;
                }

                FMainDS.ACostCentre.DefaultView.RowFilter = "";
            }
        }

        private void DisplayGLDetails()
        {
            // there is only one match?
            // TODO: support split GL transactions
            CurrentlySelectedMatch = (BankImportTDSAEpMatchRow)FMatchView[0].Row;

            if (CurrentlySelectedMatch != null)
            {
                if (CurrentlySelectedMatch.IsAccountCodeNull())
                {
                    cmbGLAccount.SelectedIndex = -1;
                }
                else
                {
                    cmbGLAccount.SetSelectedString(CurrentlySelectedMatch.AccountCode);
                }

                if (CurrentlySelectedMatch.IsCostCentreCodeNull())
                {
                    cmbGLCostCentre.SelectedIndex = -1;
                }
                else
                {
                    cmbGLCostCentre.SetSelectedString(CurrentlySelectedMatch.CostCentreCode);
                }

                if (!CurrentlySelectedMatch.IsReferenceNull())
                {
                    txtGLReference.Text = CurrentlySelectedMatch.Reference;
                }

                if (CurrentlySelectedMatch.IsNarrativeNull())
                {
                    txtGLNarrative.Text = CurrentlySelectedTransaction.Description;
                }
                else
                {
                    txtGLNarrative.Text = CurrentlySelectedMatch.Narrative;
                }
            }
        }

        private void GetGLValuesFromScreen()
        {
            if (CurrentlySelectedMatch != null)
            {
                CurrentlySelectedMatch.AccountCode = cmbGLAccount.GetSelectedString();
                CurrentlySelectedMatch.CostCentreCode = cmbGLCostCentre.GetSelectedString();
                CurrentlySelectedMatch.Reference = txtGLReference.Text;
                CurrentlySelectedMatch.Narrative = txtGLNarrative.Text;
            }
        }

        private Int32 NewMatchKey = -1;

        private void AddGiftDetail(System.Object sender, EventArgs e)
        {
            GetValuesFromScreen();

            // get a new detail number
            Int32 newDetailNumber = 0;
            decimal amount = 0;
            AEpMatchRow match = null;

            for (int i = 0; i < FMatchView.Count; i++)
            {
                match = (AEpMatchRow)FMatchView[i].Row;

                if (match.Detail >= newDetailNumber)
                {
                    newDetailNumber = match.Detail + 1;
                }

                amount += match.GiftTransactionAmount;
            }

            if (match != null)
            {
                AEpMatchRow newRow = FMainDS.AEpMatch.NewRowTyped();
                newRow.EpMatchKey = NewMatchKey--;
                newRow.MatchText = match.MatchText;
                newRow.Detail = newDetailNumber;
                newRow.LedgerNumber = match.LedgerNumber;
                newRow.AccountCode = match.AccountCode;
                newRow.CostCentreCode = match.CostCentreCode;
                newRow.DonorKey = match.DonorKey;
                newRow.GiftTransactionAmount = CurrentlySelectedTransaction.TransactionAmount -
                                               amount;
                FMainDS.AEpMatch.Rows.Add(newRow);

                // select the new gift detail
                grdGiftDetails.SelectRowInGrid(grdGiftDetails.Rows.Count-1);
                pnlEditGiftDetail.Enabled = true;
            }
        }

        private void RemoveGiftDetail(System.Object sender, EventArgs e)
        {
            GetValuesFromScreen();

            if (CurrentlySelectedMatch == null)
            {
                MessageBox.Show(Catalog.GetString("Please select a row before deleting a detail"));
                return;
            }

            // we should never allow to delete all details, otherwise we have nothing to copy from
            if (FMatchView.Count == 1)
            {
                MessageBox.Show(Catalog.GetString("At least one detail must remain"));
            }
            else
            {
                CurrentlySelectedMatch.Delete();
                CurrentlySelectedMatch = null;
                grdGiftDetails.SelectRowInGrid(-1);

                pnlEditGiftDetail.Enabled = false;
            }
        }

        private void BtnOK_Click(Object Sender, EventArgs e)
        {
            GetValuesFromScreen();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(Object Sender, EventArgs e)
        {
            //this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
    }
}