﻿//
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
    public partial class TFrmBankStatementImport
    {
        private Int32 FLedgerNumber;
        private TMBankimportNamespace FPluginRemote;

        /// <summary>
        /// use this ledger
        /// </summary>
        public Int32 LedgerNumber
        {
            set
            {
                FPluginRemote = new TMBankimportNamespace();

                FLedgerNumber = value;

                cmbBankAccount.Enabled = false;

                ALedgerRow Ledger =
                    ((ALedgerTable)TDataCache.TMFinance.GetCacheableFinanceTable(TCacheableFinanceTablesEnum.LedgerDetails, FLedgerNumber))[0];
                txtCreditSum.CurrencyCode = Ledger.BaseCurrency;
                txtDebitSum.CurrencyCode = Ledger.BaseCurrency;
                txtAmount.CurrencyCode = Ledger.BaseCurrency;

                pnlDetails.Visible = false;
            }
        }

        private DataView FMatchView = null;
        private DataView FTransactionView = null;

        private void RunOnceOnActivationManual()
        {
            TFrmSelectBankStatement DlgSelect = new TFrmSelectBankStatement(FPetraUtilsObject.GetCallerForm());

            DlgSelect.LedgerNumber = FLedgerNumber;

            if (DlgSelect.ShowDialog() == DialogResult.OK)
            {
                SelectBankStatement(DlgSelect.StatementKey);
            }
            else
            {
                // we cannot close here, because the Show method would fail. TODO: add return false?
                // this.Close();
            }
        }

        private void GetBankStatementTransactionsAndMatches(Int32 AStatementKey)
        {
            // load the transactions of the selected statement, and the matches
            FMainDS.Merge(
                FPluginRemote.WebConnectors.GetBankStatementTransactionsAndMatches(
                    AStatementKey, FLedgerNumber));
        }

        /// <summary>
        /// select the bank statement that should be loaded
        /// </summary>
        /// <param name="AStatementKey"></param>
        private void SelectBankStatement(Int32 AStatementKey)
        {
            CurrentlySelectedMatch = null;
            CurrentStatement = null;

            // merge the cost centres and the motivation details from the cacheable tables
            FMainDS.ACostCentre.Merge(TDataCache.TMFinance.GetCacheableFinanceTable(TCacheableFinanceTablesEnum.CostCentreList, FLedgerNumber));
            FMainDS.AMotivationDetail.Merge(TDataCache.TMFinance.GetCacheableFinanceTable(TCacheableFinanceTablesEnum.MotivationList, FLedgerNumber));
            FMainDS.ACostCentre.AcceptChanges();
            FMainDS.AMotivationDetail.AcceptChanges();

            // load the transactions of the selected statement, and the matches
            Thread t = new Thread(() => GetBankStatementTransactionsAndMatches(AStatementKey));

            using (TProgressDialog dialog = new TProgressDialog(t))
            {
                if (dialog.ShowDialog() == DialogResult.Cancel)
                {
                    return;
                }
            }

            while (FMainDS.AEpStatement.Rows.Count != 1)
            {
                // wait for the merging of the dataset to finish in the thread
                Thread.Sleep(300);
            }

            // an old version of the CSV import plugin did not set the potential gift typecode
            foreach (AEpTransactionRow r in FMainDS.AEpTransaction.Rows)
            {
                if (r.IsTransactionTypeCodeNull() && (r.TransactionAmount > 0))
                {
                    r.TransactionTypeCode = MFinanceConstants.BANK_STMT_POTENTIAL_GIFT;
                }
            }

            CurrentStatement = (AEpStatementRow)FMainDS.AEpStatement[0];

            FTransactionView = FMainDS.AEpTransaction.DefaultView;
            FTransactionView.AllowNew = false;
            FTransactionView.Sort = AEpTransactionTable.GetOrderDBName() + " ASC";
            grdAllTransactions.DataSource = new DevAge.ComponentModel.BoundDataView(FTransactionView);

            TFinanceControls.InitialiseMotivationGroupList(ref cmbMotivationGroup, FLedgerNumber, true);
            TFinanceControls.InitialiseMotivationDetailList(ref cmbMotivationDetail, FLedgerNumber, true);
            TFinanceControls.InitialiseCostCentreList(ref cmbGLCostCentre, FLedgerNumber, true, false, true, true);
            TFinanceControls.InitialiseAccountList(ref cmbGLAccount, FLedgerNumber, true, false, true, false);

            FMatchView = FMainDS.AEpMatch.DefaultView;
            FMatchView.AllowNew = false;
            grdGiftDetails.DataSource = new DevAge.ComponentModel.BoundDataView(FMatchView);

            TFinanceControls.InitialiseAccountList(ref cmbBankAccount, FLedgerNumber, true, false, true, true);

            if (CurrentStatement != null)
            {
                FMainDS.AEpStatement.DefaultView.RowFilter = String.Format("{0}={1}",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey);
                cmbBankAccount.SetSelectedString(CurrentStatement.BankAccountCode);
                txtBankStatement.Text = CurrentStatement.Filename;
                dtpBankStatementDate.Date = CurrentStatement.Date;
                FMainDS.AEpStatement.DefaultView.RowFilter = string.Empty;
            }

            TransactionFilterChanged(null, null);
            grdAllTransactions.SelectRowInGrid(1);
            grdAllTransactions.AutoResizeGrid();
        }

        private AEpStatementRow CurrentStatement = null;
        private BankImportTDSAEpTransactionRow CurrentlySelectedTransaction = null;
        private BankImportTDSAEpMatchRow CurrentlySelectedMatch = null;

        private void AllTransactionsFocusedRowChanged(System.Object sender, EventArgs e)
        {
            pnlDetails.Visible = true;
            pnlDetails.Enabled = false;

            CurrentlySelectedMatch = null;

            try
            {
                CurrentlySelectedTransaction = ((BankImportTDSAEpTransactionRow)grdAllTransactions.SelectedDataRowsAsDataRowView[0].Row);
            }
            catch (System.IndexOutOfRangeException)
            {
                // this can happen when the transaction type has changed, and the row disappears from the grid
                // select another row
                grdAllTransactions.SelectRowInGrid(1);
                return;
            }

            // load selections from the a_ep_match table for the new row
            FMatchView.RowFilter = AEpMatchTable.GetMatchTextDBName() +
                                   " = '" + CurrentlySelectedTransaction.MatchText + "'";

            AEpMatchRow match = (AEpMatchRow)FMatchView[0].Row;

            if (match.Action == MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT)
            {
                txtDonorKey.Text = StringHelper.FormatStrToPartnerKeyString(match.DonorKey.ToString());

                pnlGiftEdit.Visible = true;
                pnlGLEdit.Visible = false;

                grdGiftDetails.SelectRowInGrid(1);
                // grdGiftDetails.SelectRowInGrid does not seem to update the gift details, so we call that manually
                GiftDetailsFocusedRowChanged(null, null);
                grdGiftDetails.AutoResizeGrid();
                grdAllTransactions.Focus();
            }
            else if (match.Action == MFinanceConstants.BANK_STMT_STATUS_MATCHED_GL)
            {
                pnlGiftEdit.Visible = false;
                pnlGLEdit.Visible = true;

                DisplayGLDetails();
            }
            else
            {
                pnlDetails.Visible = false;
            }
        }

        private void GiftDetailsFocusedRowChanged(System.Object sender, EventArgs e)
        {
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
            }
            else
            {
                txtAmount.NumberValueDecimal = CurrentlySelectedTransaction.TransactionAmount;
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

        private void EditMatchClicked(System.Object sender, EventArgs e)
        {
            if (CurrentlySelectedTransaction != null)
            {
                TFrmMatchTransactions dlg = new TFrmMatchTransactions(this);
                dlg.MainDS = (BankImportTDS)FMainDS.Copy();
                dlg.LedgerNumber = FLedgerNumber;
                dlg.MatchText = CurrentlySelectedTransaction.MatchText;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // find all matches for this matchtext, and copy them over
                    CurrentlySelectedTransaction.MatchAction = dlg.SelectedTransaction.MatchAction;

                    DataView MatchesByText = new DataView(
                        FMainDS.AEpMatch, string.Empty,
                        AEpMatchTable.GetMatchTextDBName(),
                        DataViewRowState.CurrentRows);

                    DataRowView[] MatchesToDeleteRv = MatchesByText.FindRows(CurrentlySelectedTransaction.MatchText);

                    foreach (DataRowView rv in MatchesToDeleteRv)
                    {
                        rv.Row.Delete();
                    }

                    foreach (DataRowView rv in dlg.UpdatedMatches)
                    {
                        AEpMatchRow EditedRow = (AEpMatchRow)rv.Row;
                        AEpMatchRow NewRow = FMainDS.AEpMatch.NewRowTyped();
                        DataUtilities.CopyAllColumnValues(EditedRow, NewRow);
                        FMainDS.AEpMatch.Rows.Add(NewRow);
                    }

                    grdAllTransactions.SelectRowInGrid(1);
                    AllTransactionsFocusedRowChanged(null, null);
                }
            }
        }

        /// <summary>
        /// save the matches
        /// </summary>
        /// <returns></returns>
        public bool SaveChanges()
        {
            BankImportTDS Changes = FMainDS.GetChangesTyped(true);

            if (Changes == null)
            {
                return true;
            }

            if (FPluginRemote.WebConnectors.CommitMatches(Changes))
            {
                FMainDS.AcceptChanges();
                return true;
            }
            else
            {
                MessageBox.Show(Catalog.GetString(
                        "The matches could not be stored. Please ask your System Administrator to check the log file on the server."));
                return false;
            }
        }

        private void CreateGiftBatchThread()
        {
            TVerificationResultCollection VerificationResult;
            Int32 GiftBatchNumber = FPluginRemote.WebConnectors.CreateGiftBatch(
                FLedgerNumber,
                CurrentStatement.StatementKey,
                -1,
                out VerificationResult);

            if (GiftBatchNumber != -1)
            {
                MessageBox.Show(String.Format(Catalog.GetString("Please check Gift Batch {0}"), GiftBatchNumber));
            }
            else
            {
                if (VerificationResult != null)
                {
                    MessageBox.Show(
                        VerificationResult.BuildVerificationResultString(),
                        Catalog.GetString("Problem: No gift batch has been created"));
                }
                else
                {
                    MessageBox.Show(
                        Catalog.GetString("Unknown error"),
                        Catalog.GetString("Problem: No gift batch has been created"));
                }
            }
        }

        private void CreateGiftBatch(System.Object sender, EventArgs e)
        {
            // TODO: should we first ask? also when closing the window?
            SaveChanges();

            // load the transactions of the selected statement, and the matches
            Thread t = new Thread(() => CreateGiftBatchThread());

            using (TProgressDialog dialog = new TProgressDialog(t))
            {
                dialog.ShowDialog();
            }
        }

        private void CreateGLBatch(System.Object sender, EventArgs e)
        {
            // TODO: should we first ask? also when closing the window?
            SaveChanges();

            TVerificationResultCollection VerificationResult;
            Int32 GLBatchNumber = FPluginRemote.WebConnectors.CreateGLBatch(FMainDS,
                FLedgerNumber,
                CurrentStatement.StatementKey,
                -1,
                out VerificationResult);

            if (GLBatchNumber != -1)
            {
                MessageBox.Show(String.Format(Catalog.GetString("Please check GL Batch {0}"), GLBatchNumber));
            }
            else
            {
                if (VerificationResult != null)
                {
                    MessageBox.Show(
                        VerificationResult.BuildVerificationResultString(),
                        Catalog.GetString("Problem: No GL batch has been created"));
                }
                else
                {
                    MessageBox.Show(Catalog.GetString("Problem: No GL batch has been created"),
                        Catalog.GetString("Error"));
                }
            }
        }

        private void ExportGiftBatchThread(bool AWithInteractionOnSuccess)
        {
            TVerificationResultCollection VerificationResult;
            Int32 GiftBatchNumber = FPluginRemote.WebConnectors.CreateGiftBatch(
                FLedgerNumber,
                CurrentStatement.StatementKey,
                -1,
                out VerificationResult);

            if (GiftBatchNumber != -1)
            {
                if ((VerificationResult != null) && (VerificationResult.Count > 0))
                {
                    MessageBox.Show(
                        VerificationResult.BuildVerificationResultString(),
                        Catalog.GetString("Info: gift batch has been created"));
                }

                // export to csv
                TFrmGiftBatchExport exportForm = new TFrmGiftBatchExport(FPetraUtilsObject.GetForm());
                exportForm.LedgerNumber = FLedgerNumber;
                exportForm.FirstBatchNumber = GiftBatchNumber;
                exportForm.LastBatchNumber = GiftBatchNumber;
                exportForm.IncludeUnpostedBatches = true;

                if (TAppSettingsManager.HasValue("NumberFormat"))
                {
                    exportForm.NumberFormat = TAppSettingsManager.GetValue("NumberFormat");
                }

                exportForm.TransactionsOnly = true;
                exportForm.ExtraColumns = false;
                exportForm.OutputFilename = TAppSettingsManager.GetValue("BankImport.GiftBatchExportFilename",
                    TAppSettingsManager.GetValue("OpenPetra.PathTemp") +
                    Path.DirectorySeparatorChar +
                    "giftBatch" + GiftBatchNumber.ToString("000000") + ".csv");
                exportForm.ExportBatches(AWithInteractionOnSuccess);
            }
            else
            {
                if (VerificationResult != null)
                {
                    MessageBox.Show(
                        VerificationResult.BuildVerificationResultString(),
                        Catalog.GetString("Problem: No gift batch has been created"));
                }
                else
                {
                    MessageBox.Show(
                        Catalog.GetString("Problem: No gift batch has been created"));
                }
            }
        }

        /// <summary>
        /// this is useful for the situation, where we are using OpenPetra only for the bankimport,
        /// but need to post the gift batches in the old Petra 2.x database
        /// </summary>
        private void ExportGiftBatch(System.Object sender, EventArgs e)
        {
            ExportGiftBatch();
        }

        /// <summary>
        /// this is useful for the situation, where we are using OpenPetra only for the bankimport,
        /// but need to post the gift batches in the old Petra 2.x database
        /// </summary>
        private void ExportGiftBatch(bool AWithInteractionOnSuccess = true)
        {
            // TODO: should we first ask? also when closing the window?
            SaveChanges();

            // load the transactions of the selected statement, and the matches
            Thread t = new Thread(() => ExportGiftBatchThread(AWithInteractionOnSuccess));

            using (TProgressDialog dialog = new TProgressDialog(t))
            {
                dialog.ShowDialog();
            }
        }

        /// <summary>
        /// this exports all csv files, and all pdf files to the tmp directory specified in BankImport.GiftBatchExportPath
        /// </summary>
        private void ExportAndPrintAll(System.Object sender, EventArgs e)
        {
            string unmatchedGL = TAppSettingsManager.GetValue("BankImport.Filename.unmatched_gl", "unmatched_gl");
            string unmatched_gifts = TAppSettingsManager.GetValue("BankImport.Filename.unmatched_gifts", "unmatched_gifts");
            string matched_gifts = TAppSettingsManager.GetValue("BankImport.Filename.matched_gifts", "matched_gifts");
            string shortformat = TAppSettingsManager.GetValue("BankImport.Filename.shortformat", "shortformat");
            string all = TAppSettingsManager.GetValue("BankImport.Filename.all", "all");

            string baseFilename =
                TAppSettingsManager.GetValue("BankImport.GiftBatchExportPath") + Path.DirectorySeparatorChar +
                CurrentStatement.Filename + "_" + CurrentStatement.Date.ToString("yyyyMMdd") + "_";

            // export all transactions, as PDF
            rbtListAll.Checked = true;
            PrintReportToPDF(baseFilename + "_" + all + ".pdf",
                TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate"));

            // export matched gifts, as PDF
            rbtListGift.Checked = true;
            PrintReportToPDF(baseFilename + "_" + matched_gifts + ".pdf",
                TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate"));

            // export matched gifts, as gift batch csv file
            rbtListGift.Checked = true;
            ExportGiftBatch(false);

            if (File.Exists(baseFilename + "_giftbatch.csv"))
            {
                File.Delete(baseFilename + "_giftbatch.csv");
            }

            File.Copy(TAppSettingsManager.GetValue("BankImport.GiftBatchExportFilename"),
                baseFilename + "_giftbatch.csv");

            // export unmatched gifts, as CSV and as full pdf and as short pdf
            rbtListUnmatchedGift.Checked = true;
            ExportToExcelFile(baseFilename + "_" + unmatched_gifts + ".xlsx");
            PrintReportToPDF(baseFilename + "_" + unmatched_gifts + ".pdf",
                TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate"));
            PrintReportToPDF(baseFilename + "_" + unmatched_gifts + "_" + shortformat + ".pdf",
                TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate.ShortFormat"));

            // export unmatched GL, as CSV and as pdf
            rbtListUnmatchedGL.Checked = true;
            ExportToExcelFile(baseFilename + "_" + unmatchedGL + ".xlsx");
            PrintReportToPDF(baseFilename + "_" + unmatchedGL + ".pdf",
                TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate"));

            rbtListAll.Checked = true;

            MessageBox.Show(String.Format(Catalog.GetString("All files have been exported successfully. Please check the files {0}"), baseFilename +
                    "*"),
                Catalog.GetString("Export all pdf and csv files"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void PrintReport(System.Object sender, EventArgs e)
        {
            if (FMainDS.AEpTransaction.DefaultView.Count == 0)
            {
                return;
            }

            System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
            bool PrinterInstalled = doc.PrinterSettings.IsValid;

            if (!PrinterInstalled)
            {
                MessageBox.Show("The program cannot find a printer, and therefore cannot print!", "Problem with printing");
                return;
            }

            string letterTemplateFilename = TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate", false);

            string HtmlDocument = PrepareHTMLReport(letterTemplateFilename);

            if (HtmlDocument.Length == 0)
            {
                MessageBox.Show(Catalog.GetString("nothing to print"));
                return;
            }

            TGfxPrinter GfxPrinter = new TGfxPrinter(doc, TGfxPrinter.ePrinterBehaviour.eFormLetter);
            TPrinterHtml htmlPrinter = new TPrinterHtml(HtmlDocument,
                String.Empty,
                GfxPrinter);
            GfxPrinter.Init(eOrientation.ePortrait, htmlPrinter, eMarginType.eDefaultMargins);

            PrintDialog dlg = new PrintDialog();
            dlg.Document = GfxPrinter.Document;
            dlg.AllowCurrentPage = true;
            dlg.AllowSomePages = true;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                dlg.Document.Print();
            }
        }

        private void PrintShortReport(System.Object sender, EventArgs e)
        {
            if (FMainDS.AEpTransaction.DefaultView.Count == 0)
            {
                return;
            }

            System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
            bool PrinterInstalled = doc.PrinterSettings.IsValid;

            if (!PrinterInstalled)
            {
                MessageBox.Show("The program cannot find a printer, and therefore cannot print!", "Problem with printing");
                return;
            }

            string letterTemplateFilename = TAppSettingsManager.GetValue("BankImport.ReportHTMLTemplate.ShortFormat", false);

            string HtmlDocument = PrepareHTMLReport(letterTemplateFilename);

            if (HtmlDocument.Length == 0)
            {
                MessageBox.Show(Catalog.GetString("nothing to print"));
                return;
            }

            TGfxPrinter GfxPrinter = new TGfxPrinter(doc, TGfxPrinter.ePrinterBehaviour.eFormLetter);
            TPrinterHtml htmlPrinter = new TPrinterHtml(HtmlDocument,
                String.Empty,
                GfxPrinter);
            GfxPrinter.Init(eOrientation.ePortrait, htmlPrinter, eMarginType.eDefaultMargins);

            PrintDialog dlg = new PrintDialog();
            dlg.Document = GfxPrinter.Document;
            dlg.AllowCurrentPage = true;
            dlg.AllowSomePages = true;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                dlg.Document.Print();
            }
        }

        private void PrintReportToPDF(string ASaveASFilename, string ALetterTemplateFilename)
        {
            if (FMainDS.AEpTransaction.DefaultView.Count == 0)
            {
                return;
            }

            string HtmlDocument = PrepareHTMLReport(ALetterTemplateFilename);

            if (HtmlDocument.Length == 0)
            {
                return;
            }

            // print to pdf
            PrintDocument doc = new PrintDocument();

            TPdfPrinter pdfPrinter = new TPdfPrinter(doc, TGfxPrinter.ePrinterBehaviour.eFormLetter);
            TPrinterHtml htmlPrinter = new TPrinterHtml(HtmlDocument,
                String.Empty,
                pdfPrinter);
            pdfPrinter.Init(eOrientation.ePortrait, htmlPrinter, eMarginType.ePrintableArea);

            pdfPrinter.SavePDF(ASaveASFilename);
        }

        private string PrepareHTMLReport(string ALetterTemplateFilename)
        {
            string ShortCodeOfBank = txtBankStatement.Text;
            string DateOfStatement = StringHelper.DateToLocalizedString(
                dtpBankStatementDate.Date.HasValue ? dtpBankStatementDate.Date.Value : DateTime.Today);
            string HtmlDocument = String.Empty;

            if (rbtListAll.Checked)
            {
                HtmlDocument =
                    PrintHTML(
                        CurrentStatement,
                        FMainDS.AEpTransaction.DefaultView, FMainDS.AEpMatch, Catalog.GetString(
                            "Full bank statement") + ", " + ShortCodeOfBank + ", " + DateOfStatement,
                        ALetterTemplateFilename);
            }
            else if (rbtListUnmatchedGift.Checked)
            {
                HtmlDocument =
                    PrintHTML(
                        CurrentStatement,
                        FMainDS.AEpTransaction.DefaultView, FMainDS.AEpMatch, Catalog.GetString(
                            "Unmatched gifts") + ", " + ShortCodeOfBank + ", " + DateOfStatement,
                        ALetterTemplateFilename);
            }
            else if (rbtListUnmatchedGL.Checked)
            {
                HtmlDocument =
                    PrintHTML(
                        CurrentStatement,
                        FMainDS.AEpTransaction.DefaultView, FMainDS.AEpMatch, Catalog.GetString(
                            "Unmatched GL") + ", " + ShortCodeOfBank + ", " + DateOfStatement,
                        ALetterTemplateFilename);
            }
            else if (rbtListGift.Checked)
            {
                HtmlDocument =
                    PrintHTML(
                        CurrentStatement,
                        FMainDS.AEpTransaction.DefaultView, FMainDS.AEpMatch, Catalog.GetString(
                            "Matched gifts") + ", " + ShortCodeOfBank + ", " + DateOfStatement,
                        ALetterTemplateFilename);
            }

            return HtmlDocument;
        }

        private static string RemoveSEPAText(string ADescription)
        {
            if (ADescription.Contains("EREF+ZV") && ADescription.Contains("PURP+RINP"))
            {
                ADescription = ADescription.Substring(0, ADescription.IndexOf("EREF+ZV")) + ADescription.Substring(ADescription.IndexOf("PURP+RINP"));
            }

            if (ADescription.Contains("IBAN:") && ADescription.Contains("BIC:"))
            {
                ADescription = ADescription.Substring(0, ADescription.IndexOf("IBAN:"));

                if (ADescription.Contains("EREF:"))
                {
                    ADescription = ADescription.Substring(0, ADescription.IndexOf("EREF:"));
                }

                if (ADescription.Contains("EREF:"))
                {
                    ADescription = ADescription.Substring(0, ADescription.IndexOf("EREF:"));
                }
            }

            ADescription = ADescription.Replace("PURP+RINPRATENZAHLUNG", "");

            if (ADescription.Contains("ABWA+"))
            {
                // abweichender Zahlungsauftraggeber
                ADescription = ADescription.Substring(0, ADescription.IndexOf("ABWA+"));
            }

            ADescription = ADescription.Replace("SVWZ+", "");

            return ADescription.Trim();
        }

        /// <summary>
        /// dump unmatched gifts or other transactions to a HTML table for printing
        /// </summary>
        private static string PrintHTML(
            AEpStatementRow ACurrentStatement,
            DataView AEpTransactions, AEpMatchTable AMatches, string ATitle,
            string ALetterTemplateFilename)
        {
            if ((ALetterTemplateFilename.Length == 0) || !File.Exists(ALetterTemplateFilename))
            {
                OpenFileDialog DialogOpen = new OpenFileDialog();
                DialogOpen.Filter = "Report template (*.html)|*.html";
                DialogOpen.RestoreDirectory = true;
                DialogOpen.Title = "Open Report Template";

                if (DialogOpen.ShowDialog() == DialogResult.OK)
                {
                    ALetterTemplateFilename = DialogOpen.FileName;
                }
            }

            // message body from HTML template
            StreamReader reader = new StreamReader(ALetterTemplateFilename);
            string msg = reader.ReadToEnd();

            reader.Close();

            msg = msg.Replace("#TITLE", ATitle);
            msg = msg.Replace("#PRINTDATE", DateTime.Now.ToShortDateString());

            if (!ACurrentStatement.IsIdFromBankNull())
            {
                msg = msg.Replace("#STATEMENTNR", ACurrentStatement.IdFromBank);
            }

            if (!ACurrentStatement.IsStartBalanceNull())
            {
                msg = msg.Replace("#STARTBALANCE", String.Format("{0:N}", ACurrentStatement.StartBalance));
            }

            if (!ACurrentStatement.IsEndBalanceNull())
            {
                msg = msg.Replace("#ENDBALANCE", String.Format("{0:N}", ACurrentStatement.EndBalance));
            }

            // recognise detail lines automatically
            string RowTemplate;
            msg = TPrinterHtml.GetTableRow(msg, "#NRONSTATEMENT", out RowTemplate);
            string rowTexts = "";

            BankImportTDSAEpTransactionRow row = null;

            AEpTransactions.Sort = BankImportTDSAEpTransactionTable.GetNumberOnPaperStatementDBName();

            Decimal Sum = 0.0m;
            Int32 NumberPrinted = 0;

            DataView MatchesByMatchText = new DataView(AMatches,
                string.Empty,
                AEpMatchTable.GetMatchTextDBName(),
                DataViewRowState.CurrentRows);

            string thinLine = "<font size=\"-3\">-------------------------------------------------------------------------<br/></font>";

            foreach (DataRowView rv in AEpTransactions)
            {
                row = (BankImportTDSAEpTransactionRow)rv.Row;

                string rowToPrint = RowTemplate;

                // short description, remove all SEPA stuff
                string ShortDescription = RemoveSEPAText(row.Description);

                rowToPrint = rowToPrint.Replace("#NAME", row.AccountName);
                rowToPrint = rowToPrint.Replace("#DESCRIPTION", row.Description);
                rowToPrint = rowToPrint.Replace("#SHORTDESCRIPTION", ShortDescription);

                string RecipientDescription = string.Empty;

                DataRowView[] matches = MatchesByMatchText.FindRows(row.MatchText);

                AEpMatchRow match = null;

                foreach (DataRowView rvMatch in matches)
                {
                    match = (AEpMatchRow)rvMatch.Row;

                    if (RecipientDescription.Length > 0)
                    {
                        RecipientDescription += "<br/>";
                    }

                    if (!match.IsRecipientKeyNull() && (match.RecipientKey > 0))
                    {
                        RecipientDescription += match.RecipientKey.ToString() + " ";
                    }

                    RecipientDescription += match.RecipientShortName;
                }

                if (RecipientDescription.Trim().Length > 0)
                {
                    rowToPrint = rowToPrint.Replace("#RECIPIENTDESCRIPTIONUNMATCHED", string.Empty);
                    rowToPrint = rowToPrint.Replace("#RECIPIENTDESCRIPTION", "<br/>" + thinLine + RecipientDescription);
                }
                else
                {
                    // extra space for unmatched gifts
                    rowToPrint = rowToPrint.Replace("#RECIPIENTDESCRIPTIONUNMATCHED", "<br/><br/>");
                    rowToPrint = rowToPrint.Replace("#RECIPIENTDESCRIPTION", string.Empty);
                }

                if ((match != null) && !match.IsDonorKeyNull() && (match.DonorKey > 0))
                {
                    string DonorDescription = "<br/>" + thinLine + match.DonorKey.ToString() + " " + match.DonorShortName;

                    rowToPrint = rowToPrint.Replace("#DONORDESCRIPTION", DonorDescription);
                    rowToPrint = rowToPrint.Replace("#DONORKEY", StringHelper.PartnerKeyToStr(match.DonorKey));
                    rowToPrint = rowToPrint.Replace("#DONORNAMEORDESCRIPTION", match.DonorShortName);
                }
                else
                {
                    rowToPrint = rowToPrint.Replace("#DONORDESCRIPTION", string.Empty);
                    rowToPrint = rowToPrint.Replace("#DONORKEY", string.Empty);
                    rowToPrint = rowToPrint.Replace("#DONORNAMEORDESCRIPTION", row.AccountName);
                }

                rowTexts += rowToPrint.
                            Replace("#NRONSTATEMENT", row.NumberOnPaperStatement.ToString()).
                            Replace("#AMOUNT", String.Format("{0:C}", row.TransactionAmount)).
                            Replace("#IBANANDBIC",
                    row.IsIbanNull() ? string.Empty : "<br/>" + row.Iban + "<br/>" + row.Bic).
                            Replace("#IBAN", row.Iban).
                            Replace("#BIC", row.Bic).
                            Replace("#ACCOUNTNUMBER", row.BankAccountNumber).
                            Replace("#BANKSORTCODE", row.BranchCode);

                Sum += Convert.ToDecimal(row.TransactionAmount);
                NumberPrinted++;
            }

            Sum = Math.Round(Sum, 2);

            msg = msg.Replace("#ROWTEMPLATE", rowTexts);
            msg = msg.Replace("#TOTALAMOUNT", String.Format("{0:C}", Sum));
            msg = msg.Replace("#TOTALNUMBER", NumberPrinted.ToString());

            return msg;
        }

        private bool ExportToExcelFile(string AFilename)
        {
            BankImportTDSAEpTransactionRow row = null;

            FMainDS.AEpTransaction.DefaultView.Sort = BankImportTDSAEpTransactionTable.GetNumberOnPaperStatementDBName();

            DataView MatchesByMatchText = new DataView(FMainDS.AEpMatch,
                string.Empty,
                AEpMatchTable.GetMatchTextDBName(),
                DataViewRowState.CurrentRows);

            List <string>Lines = new List <string>();
            Lines.Add("Order,AccountName,DonorKey,DonorName,IBAN,BIC,AccountNumber,BankCode,Description,Amount");

            if (FMainDS.AEpTransaction.DefaultView.Count == 0)
            {
                if (File.Exists(AFilename))
                {
                    File.Delete(AFilename);
                }

                return true;
            }

            foreach (DataRowView rv in FMainDS.AEpTransaction.DefaultView)
            {
                row = (BankImportTDSAEpTransactionRow)rv.Row;

                string line = string.Empty;
                line = StringHelper.AddCSV(line, new TVariant(row.NumberOnPaperStatement).EncodeToString());
                line = StringHelper.AddCSV(line, row.AccountName);

                DataRowView[] matches = MatchesByMatchText.FindRows(row.MatchText);

                AEpMatchRow match = null;

                foreach (DataRowView rvMatch in matches)
                {
                    match = (AEpMatchRow)rvMatch.Row;
                    break;
                }

                if ((match != null) && !match.IsDonorKeyNull() && (match.DonorKey > 0))
                {
                    line = StringHelper.AddCSV(line, StringHelper.PartnerKeyToStr(match.DonorKey));
                    line = StringHelper.AddCSV(line, match.DonorShortName);
                }
                else
                {
                    line = StringHelper.AddCSV(line, string.Empty); // donorkey
                    line = StringHelper.AddCSV(line, string.Empty); // donordescription
                }

                line = StringHelper.AddCSV(line, row.Iban);
                line = StringHelper.AddCSV(line, row.Bic);
                line = StringHelper.AddCSV(line, new TVariant(row.BankAccountNumber).EncodeToString());
                line = StringHelper.AddCSV(line, new TVariant(row.BranchCode).EncodeToString());
                line = StringHelper.AddCSV(line, row.Description);
                line = StringHelper.AddCSV(line, new TVariant(row.TransactionAmount).EncodeToString());
                Lines.Add(line);
            }

            XmlDocument doc = TCsv2Xml.ParseCSV2Xml(Lines, ",");

            if (doc != null)
            {
                using (FileStream fs = new FileStream(AFilename, FileMode.Create))
                {
                    if (TCsv2Xml.Xml2ExcelStream(doc, fs, false))
                    {
                        fs.Close();
                    }
                }

                return true;
            }

            return false;
        }

        private void TransactionFilterChanged(System.Object sender, EventArgs e)
        {
            pnlDetails.Visible = false;
            CurrentlySelectedMatch = null;
            CurrentlySelectedTransaction = null;

            if (FTransactionView == null)
            {
                return;
            }

            if (rbtListAll.Checked)
            {
                FTransactionView.RowFilter = String.Format("{0}={1}",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey);
            }
            else if (rbtListGift.Checked)
            {
                // TODO: allow splitting a transaction, one part is GL/AP, the other is a donation?
                //       at Top Level: split transaction, results into 2 rows in aeptransaction (not stored). Merge Transactions again?

                FTransactionView.RowFilter = String.Format("{0}={1} and {2}='{3}'",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey,
                    BankImportTDSAEpTransactionTable.GetMatchActionDBName(),
                    MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT);
            }
            else if (rbtListUnmatchedGift.Checked)
            {
                FTransactionView.RowFilter = String.Format("{0}={1} and {2}='{3}' and {4} LIKE '%{5}'",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey,
                    BankImportTDSAEpTransactionTable.GetMatchActionDBName(),
                    MFinanceConstants.BANK_STMT_STATUS_UNMATCHED,
                    BankImportTDSAEpTransactionTable.GetTransactionTypeCodeDBName(),
                    MFinanceConstants.BANK_STMT_POTENTIAL_GIFT);
            }
            else if (rbtListUnmatchedGL.Checked)
            {
                FTransactionView.RowFilter = String.Format("{0}={1} and {2}='{3}' and ({4} NOT LIKE '%{5}' OR {4} IS NULL)",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey,
                    BankImportTDSAEpTransactionTable.GetMatchActionDBName(),
                    MFinanceConstants.BANK_STMT_STATUS_UNMATCHED,
                    BankImportTDSAEpTransactionTable.GetTransactionTypeCodeDBName(),
                    MFinanceConstants.BANK_STMT_POTENTIAL_GIFT);
            }
            else if (rbtListIgnored.Checked)
            {
                FTransactionView.RowFilter = String.Format("{0}={1} and {2}='{3}'",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey,
                    BankImportTDSAEpTransactionTable.GetMatchActionDBName(),
                    MFinanceConstants.BANK_STMT_STATUS_NO_MATCHING);
            }
            else if (rbtListGL.Checked)
            {
                FTransactionView.RowFilter = String.Format("{0}={1} and {2}='{3}'",
                    AEpStatementTable.GetStatementKeyDBName(),
                    CurrentStatement.StatementKey,
                    BankImportTDSAEpTransactionTable.GetMatchActionDBName(),
                    MFinanceConstants.BANK_STMT_STATUS_MATCHED_GL);
            }

            // update sumcredit and sumdebit
            decimal sumCredit = 0.0M;
            decimal sumDebit = 0.0M;

            foreach (DataRowView rv in FTransactionView)
            {
                AEpTransactionRow Row = (AEpTransactionRow)rv.Row;

                if (Row.TransactionAmount < 0)
                {
                    sumDebit += Row.TransactionAmount * -1.0M;
                }
                else
                {
                    sumCredit += Row.TransactionAmount;
                }
            }

            txtCreditSum.NumberValueDecimal = sumCredit;
            txtDebitSum.NumberValueDecimal = sumDebit;
            txtTransactionCount.Text = FTransactionView.Count.ToString();

            if (FTransactionView.Count > 0)
            {
                grdAllTransactions.SelectRowInGrid(1);
                AllTransactionsFocusedRowChanged(null, null);
            }
        }

        /// <summary>
        /// Get the number of changed records and specify a message to incorporate into the 'Do you want to save?' message box
        /// </summary>
        /// <param name="AMessage">An optional message to display. If the parameter is an empty string a default message will be used</param>
        /// <returns>The number of changed records. Return -1 to imply 'unknown'.</returns>
        public int GetChangedRecordCount(out string AMessage)
        {
            AMessage = String.Empty;
            return -1;
        }
    }
}