//
// DO NOT REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
//
// @Authors:
//       timop
//
// Copyright 2004-2017 by OM International
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
using System.Data.Odbc;
using System.IO;
using System.Xml;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Web;
using GNU.Gettext;
using Ict.Common;
using Ict.Common.Verification;
using Ict.Common.Data; // Needed indirectly by Ict.Petra.Server.lib.MFinance.Common.dll and Ict.Petra.Shared.lib.data.dll
using Ict.Common.DB;
using Ict.Common.Remoting.Server;
using Ict.Common.Session;
using Ict.Petra.Shared;
using Ict.Petra.Shared.MFinance.GL.Data;
using Ict.Petra.Shared.MFinance.Account.Data;
using Ict.Petra.Shared.MFinance.Gift.Data;
using Ict.Petra.Shared.MPartner.Partner.Data;
using Ict.Petra.Shared.MPartner;
using Ict.Petra.Shared.MFinance;
using Ict.Petra.Server.MFinance.Cacheable;
using Ict.Petra.Server.MFinance.Account.Data.Access;
using Ict.Petra.Server.MFinance.Gift.Data.Access;
using Ict.Petra.Server.MPartner.Partner.Data.Access;
using Ict.Petra.Plugins.Bankimport.Data;
using Ict.Petra.Plugins.Bankimport.Data.Access;
using Ict.Petra.Server.MFinance.GL;
using Ict.Petra.Server.App.Core.Security;
using Ict.Petra.Server.MFinance.Common;
using Ict.Petra.Server.MFinance.Gift.WebConnectors;
using Ict.Petra.Server.MFinance.GL.WebConnectors;
using Ict.Petra.Server.App.Core;
using Ict.Petra.Plugins.Bankimport.Server;

namespace Ict.Petra.Plugins.Bankimport.WebConnectors
{
    /// <summary>
    /// import a bank statement from a CSV file
    /// </summary>
    public class TBankImportWebConnector
    {
        /// <summary>
        /// upload new bank statement so that it can be used for matching etc.
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static TSubmitChangesResult StoreNewBankStatement(BankImportTDS AStatementAndTransactionsDS,
            out Int32 AFirstStatementKey)
        {
            string MyClientID = DomainManager.GClientID.ToString();

            AFirstStatementKey = -1;

            TProgressTracker.InitProgressTracker(MyClientID,
                Catalog.GetString("Processing new bank statements"),
                AStatementAndTransactionsDS.AEpStatement.Rows.Count + 1);

            TProgressTracker.SetCurrentState(MyClientID,
                Catalog.GetString("Saving to database"),
                0);

            try
            {
                // Must not throw away the changes because we need the correct statement keys
                AStatementAndTransactionsDS.DontThrowAwayAfterSubmitChanges = true;
                BankImportTDSAccess.SubmitChanges(AStatementAndTransactionsDS);

                AFirstStatementKey = -1;

                if (AStatementAndTransactionsDS != null)
                {
                    TProgressTracker.SetCurrentState(MyClientID,
                        Catalog.GetString("starting to train"),
                        1);

                    AFirstStatementKey = AStatementAndTransactionsDS.AEpStatement[0].StatementKey;

                    // search for already posted gift batches, and do the matching for these imported statements
                    TBankImportMatching.Train(AStatementAndTransactionsDS.AEpStatement);
                }

                TProgressTracker.FinishJob(MyClientID);
            }
            catch (Exception ex)
            {
                TLogging.Log(ex.ToString());
                TProgressTracker.CancelJob(MyClientID);
                return TSubmitChangesResult.scrError;
            }

            return TSubmitChangesResult.scrOK;
        }

        /// <summary>
        /// train an existing bank statement, where the gift batch has been posted already
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static bool TrainBankStatement(Int32 ALedgerNumber, DateTime ADateOfStatement, string ABankAccountCode)
        {
            // get the statement keys
            TDBTransaction ReadTransaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadCommitted);

            AEpStatementTable Statements = new AEpStatementTable();
            AEpStatementRow row = Statements.NewRowTyped(false);

            row.LedgerNumber = ALedgerNumber;
            row.Date = ADateOfStatement;
            row.BankAccountCode = ABankAccountCode;

            Statements = AEpStatementAccess.LoadUsingTemplate(row, ReadTransaction);

            DBAccess.GDBAccessObj.RollbackTransaction();

            if (Statements.Rows.Count == 0)
            {
                return false;
            }

            // search for already posted gift batches, and do the matching for these imported statements
            TBankImportMatching.Train(Statements);

            return true;
        }

        /// <summary>
        /// train all bank statements of last month, in this ledger
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static bool TrainBankStatementsLastMonth(Int32 ALedgerNumber, DateTime AToday)
        {
            string sessionID = TSession.GetSessionID();
            Int32 clientID = DomainManager.GClientID;
            Thread t = new Thread(() => TrainBankStatementsLastMonthThread(
                    sessionID,
                    clientID,
                    ALedgerNumber, AToday));

            t.Name = Guid.NewGuid().ToString();

            t.Start();
            return true;
        }

        private static bool TrainBankStatementsLastMonthThread(String ASessionID, Int32 AClientID, Int32 ALedgerNumber, DateTime AToday)
        {
            TSession.InitThread(ASessionID);

            DomainManager.GClientID = AClientID;
            string MyClientID = DomainManager.GClientID.ToString();

            TProgressTracker.InitProgressTracker(MyClientID,
                Catalog.GetString("Training last months bank statements"),
                30);

            DateTime startDateThisMonth = new DateTime(AToday.Year, AToday.Month, 1);
            DateTime endDateLastMonth = startDateThisMonth.AddDays(-1);
            DateTime startDateLastMonth = new DateTime(endDateLastMonth.Year, endDateLastMonth.Month, 1);

            // for debugging the training:
            //startDateLastMonth = new DateTime(2015,10,5);
            //startDateThisMonth = startDateLastMonth.AddDays(1);

            // get all bank accounts
            TCacheable CachePopulator = new TCacheable();
            Type typeofTable;
            GLSetupTDSAAccountTable accounts = (GLSetupTDSAAccountTable)CachePopulator.GetCacheableTable(
                TCacheableFinanceTablesEnum.AccountList,
                "",
                false,
                ALedgerNumber,
                out typeofTable);

            foreach (GLSetupTDSAAccountRow account in accounts.Rows)
            {
                // at OM Germany we don't have the bank account flags set
                if ((!account.IsBankAccountFlagNull() && (account.BankAccountFlag == true))
                    || (!account.IsCashAccountFlagNull() && (account.CashAccountFlag == true))
                    )
                {
                    string BankAccountCode = account.AccountCode;

                    DateTime counter = startDateLastMonth;

                    while (!counter.Equals(startDateThisMonth))
                    {
                        TProgressTracker.SetCurrentState(
                            MyClientID,
                            String.Format(Catalog.GetString("Training {0} {1}"), BankAccountCode, counter.ToShortDateString()),
                            counter.Day);

                        // TODO: train only one bank statement per date and bank account?
                        TrainBankStatement(ALedgerNumber, counter, BankAccountCode);
                        counter = counter.AddDays(1);
                    }
                }
            }

            TProgressTracker.FinishJob(MyClientID);

            return true;
        }

        /// <summary>
        /// returns the bank statements that are from or newer than the given date
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static AEpStatementTable GetImportedBankStatements(Int32 ALedgerNumber, DateTime AStartDate)
        {
            TDBTransaction ReadTransaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadCommitted);

            AEpStatementTable localTable = new AEpStatementTable();
            AEpStatementRow row = localTable.NewRowTyped(false);

            row.LedgerNumber = ALedgerNumber;
            row.Date = AStartDate;

            StringCollection operators = new StringCollection();
            operators.Add("=");
            operators.Add(">=");

            localTable = AEpStatementAccess.LoadUsingTemplate(row, operators, null, ReadTransaction);

            DBAccess.GDBAccessObj.RollbackTransaction();

            return localTable;
        }

        /// <summary>
        /// drop a bank statement and all its transactions
        /// </summary>
        /// <param name="AEpStatementKey"></param>
        /// <returns></returns>
        [RequireModulePermission("FINANCE-1")]
        public static bool DropBankStatement(Int32 AEpStatementKey)
        {
            TDBTransaction Transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadCommitted);

            BankImportTDS MainDS = new BankImportTDS();

            AEpStatementAccess.LoadByPrimaryKey(MainDS, AEpStatementKey, Transaction);
            AEpTransactionAccess.LoadViaAEpStatement(MainDS, AEpStatementKey, Transaction);

            DBAccess.GDBAccessObj.RollbackTransaction();

            foreach (AEpStatementRow stmtRow in MainDS.AEpStatement.Rows)
            {
                stmtRow.Delete();
            }

            foreach (AEpTransactionRow transactionRow in MainDS.AEpTransaction.Rows)
            {
                transactionRow.Delete();
            }

            MainDS.ThrowAwayAfterSubmitChanges = true;
            try
            {
                BankImportTDSAccess.SubmitChanges(MainDS);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool FindDonorByAccountNumber(
            DataView APartnerByBankAccount,
            string ABankSortCode,
            string AAccountNumber,
            out string ADonorShortName,
            out Int64 ADonorKey)
        {
            ADonorShortName = String.Empty;
            ADonorKey = -1;

            DataRowView[] rows = APartnerByBankAccount.FindRows(new object[] { ABankSortCode, AAccountNumber });

            if (rows.Length == 1)
            {
                ADonorShortName = rows[0].Row["ShortName"].ToString();
                ADonorKey = Convert.ToInt64(rows[0].Row["PartnerKey"]);
                return true;
            }

            return false;
        }

        private struct MatchDate
        {
            public MatchDate(AEpMatchRow AR, DateTime AD)
            {
                r = AR;
                d = AD;
            }

            public AEpMatchRow r;
            public DateTime d;
        }

        private struct MatchDonor
        {
            public MatchDonor(AEpMatchRow AR, string ABankCode, string AAccountCode)
            {
                r = AR;
                a = AAccountCode;
                b = ABankCode; 
            }

            public AEpMatchRow r;
            public string a, b;
        }

        /// <summary>
        /// returns the transactions of the bank statement, and the matches if they exist;
        /// tries to find matches too
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static BankImportTDS GetBankStatementTransactionsAndMatches(Int32 AStatementKey, Int32 ALedgerNumber)
        {
            TDBTransaction Transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.Serializable);

            BankImportTDS ResultDataset = new BankImportTDS();
            string MyClientID = DomainManager.GClientID.ToString();

            TProgressTracker.InitProgressTracker(MyClientID,
                Catalog.GetString("Load Bank Statement"),
                100.0m);

            TProgressTracker.SetCurrentState(MyClientID,
                Catalog.GetString("loading statement"),
                0);

            try
            {
                AEpStatementAccess.LoadByPrimaryKey(ResultDataset, AStatementKey, Transaction);

                if (ResultDataset.AEpStatement[0].BankAccountCode.Length == 0)
                {
                    throw new Exception("Loading Bank Statement: Bank Account must not be empty");
                }

                ACostCentreAccess.LoadViaALedger(ResultDataset, ALedgerNumber, Transaction);

                AMotivationDetailAccess.LoadViaALedger(ResultDataset, ALedgerNumber, Transaction);

                AEpTransactionAccess.LoadViaAEpStatement(ResultDataset, AStatementKey, Transaction);

                BankImportTDS TempDataset = new BankImportTDS();
                AEpTransactionAccess.LoadViaAEpStatement(TempDataset, AStatementKey, Transaction);
                AEpMatchAccess.LoadViaALedger(TempDataset, ResultDataset.AEpStatement[0].LedgerNumber, Transaction);

                // load all bankingdetails and partner shortnames related to this statement
                string sqlLoadPartnerByBankAccount =
                    "SELECT DISTINCT p.p_partner_key_n AS PartnerKey, " +
                    "p.p_partner_short_name_c AS ShortName, " +
                    "t.p_branch_code_c AS BranchCode, " +
                    "t.a_bank_account_number_c AS BankAccountNumber " +
                    "FROM PUB_a_ep_transaction t, PUB_p_banking_details bd, PUB_p_bank b, PUB_p_partner_banking_details pbd, PUB_p_partner p " +
                    "WHERE t.a_statement_key_i = " + AStatementKey.ToString() + " " +
                    "AND bd.p_bank_account_number_c = t.a_bank_account_number_c " +
                    "AND b.p_partner_key_n = bd.p_bank_key_n " +
                    "AND b.p_branch_code_c = t.p_branch_code_c " +
                    "AND pbd.p_banking_details_key_i = bd.p_banking_details_key_i " +
                    "AND p.p_partner_key_n = pbd.p_partner_key_n";

                DataTable PartnerByBankAccount = DBAccess.GDBAccessObj.SelectDT(sqlLoadPartnerByBankAccount, "partnerByBankAccount", Transaction);
                PartnerByBankAccount.DefaultView.Sort = "BranchCode, BankAccountNumber";

                // get all recipients that have been merged
                string sqlGetMergedRecipients =
                    string.Format(
                        "SELECT DISTINCT p.p_partner_key_n AS PartnerKey, p.p_status_code_c AS StatusCode FROM PUB_a_ep_match m, PUB_p_partner p " +
                        "WHERE (m.p_recipient_key_n = p.p_partner_key_n OR m.p_donor_key_n = p.p_partner_key_n) AND p.p_status_code_c = '{0}'",
                        MPartnerConstants.PARTNERSTATUS_MERGED);
                DataTable MergedPartners = DBAccess.GDBAccessObj.SelectDT(sqlGetMergedRecipients, "mergedPartners", Transaction);
                MergedPartners.DefaultView.Sort = "PartnerKey";

                // get all recipients that are not active anymore
                string sqlGetNotActiveRecipients =
                    string.Format(
                        "SELECT DISTINCT p.p_partner_key_n AS PartnerKey, p.p_status_code_c AS StatusCode FROM PUB_a_ep_match m, PUB_p_partner p " +
                        "WHERE m.p_recipient_key_n = p.p_partner_key_n AND p.p_status_code_c <> '{0}'",
                        MPartnerConstants.PARTNERSTATUS_ACTIVE);
                DataTable NotActiveRecipients = DBAccess.GDBAccessObj.SelectDT(sqlGetNotActiveRecipients, "notActiveRecipients", Transaction);
                NotActiveRecipients.DefaultView.Sort = "PartnerKey";

                DBAccess.GDBAccessObj.RollbackTransaction();

                string BankAccountCode = ResultDataset.AEpStatement[0].BankAccountCode;

                TempDataset.AEpMatch.DefaultView.Sort = AEpMatchTable.GetMatchTextDBName();

                SortedList <string, AEpMatchRow>MatchesToAddLater = new SortedList <string, AEpMatchRow>();
                List <MatchDate>NewDates = new List <MatchDate>();
                List <AEpMatchRow>SetUnmatched = new List <AEpMatchRow>();
                List <MatchDonor>FindDonorKey = new List<MatchDonor>();

                int count = 0;

                // load the matches or create new matches
                foreach (BankImportTDSAEpTransactionRow row in ResultDataset.AEpTransaction.Rows)
                {
                    TProgressTracker.SetCurrentState(MyClientID,
                        Catalog.GetString("finding matches") +
                        " " + count + "/" + ResultDataset.AEpTransaction.Rows.Count.ToString(),
                        10.0m + (count * 80.0m / ResultDataset.AEpTransaction.Rows.Count));
                    count++;

                    BankImportTDSAEpTransactionRow tempTransactionRow =
                        (BankImportTDSAEpTransactionRow)TempDataset.AEpTransaction.Rows.Find(
                            new object[] {
                                row.StatementKey,
                                row.Order,
                                row.DetailKey
                            });

                    // find a match with the same match text, or create a new one
                    if (row.IsMatchTextNull() || (row.MatchText.Length == 0) || !row.MatchText.StartsWith(BankAccountCode))
                    {
                        row.MatchText = TBankImportMatching.CalculateMatchText(BankAccountCode, row);

                        tempTransactionRow.MatchText = row.MatchText;
                    }

                    DataRowView[] matches = TempDataset.AEpMatch.DefaultView.FindRows(row.MatchText);

                    if (matches.Length > 0)
                    {
                        Decimal sum = 0.0m;

                        // update the recent date
                        foreach (DataRowView rv in matches)
                        {
                            AEpMatchRow r = (AEpMatchRow)rv.Row;

                            sum += r.GiftTransactionAmount;
                            string action = r.Action;

                            // check if the recipient key is still valid. could be that they have married, and merged into another family record
                            if ((r.RecipientKey != 0)
                                && (MergedPartners.DefaultView.FindRows(r.RecipientKey).Length > 0))
                            {
                                TLogging.LogAtLevel(1, "partner has been merged: " + r.RecipientKey.ToString());
                                r.RecipientKey = 0;
                                action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
                            }

                            // check if the donor key is still valid. could be that they have married, and merged into another family record
                            if ((r.DonorKey != 0)
                                && (MergedPartners.DefaultView.FindRows(r.DonorKey).Length > 0))
                            {
                                TLogging.LogAtLevel(1, "partner has been merged: " + r.DonorKey.ToString());
                                r.DonorKey = 0;
                                action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
                            }

                            // check if the recipient is still active
                            if ((r.RecipientKey != 0)
                                && (NotActiveRecipients.DefaultView.FindRows(r.RecipientKey).Length > 0))
                            {
                                r.RecipientKey = 0;
                                action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
                            }

                            // check if the costcentre is still active
                            ACostCentreRow costcentre = (ACostCentreRow)ResultDataset.ACostCentre.Rows.Find(new object[] { ALedgerNumber,
                                                                                                                           r.CostCentreCode });

                            if ((costcentre != null) && !costcentre.CostCentreActiveFlag)
                            {
                                TLogging.LogAtLevel(1, "costcentre " + r.CostCentreCode + " is not active anymore; donor: " + r.DonorKey.ToString());
                                action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
                            }

                            if (r.RecentMatch < row.DateEffective)
                            {
                                // do not modify RecentMatch here for speed reasons
                                // r.RecentMatch = row.DateEffective;
                                NewDates.Add(new MatchDate(r, row.DateEffective));
                            }

                            if (action == MFinanceConstants.BANK_STMT_STATUS_UNMATCHED && r.Action != action)
                            {
                                SetUnmatched.Add(r);
                            }

                            row.MatchAction = action;

                            if (r.IsDonorKeyNull() || (r.DonorKey <= 0))
                            {
                                FindDonorKey.Add(new MatchDonor(r, row.BranchCode, row.BankAccountNumber));
                            }
                        }

                        if (sum != row.TransactionAmount)
                        {
                            TLogging.Log(
                                "we should drop this match since the total is wrong: " + row.Description + " " + sum.ToString() + " " +
                                row.TransactionAmount.ToString());
                            row.MatchAction = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;

                            foreach (DataRowView rv in matches)
                            {
                                AEpMatchRow r = (AEpMatchRow)rv.Row;

                                if (!SetUnmatched.Contains(r))
                                {
                                    SetUnmatched.Add(r);
                                }
                            }
                        }
                    }
                    else if (!MatchesToAddLater.ContainsKey(row.MatchText))
                    {
                        // create new match
                        AEpMatchRow tempRow = TempDataset.AEpMatch.NewRowTyped(true);
                        tempRow.EpMatchKey = (TempDataset.AEpMatch.Count + MatchesToAddLater.Count + 1) * -1;
                        tempRow.Detail = 0;
                        tempRow.MatchText = row.MatchText;
                        tempRow.LedgerNumber = ALedgerNumber;
                        tempRow.GiftTransactionAmount = row.TransactionAmount;
                        tempRow.Action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;

                        String DonorShortName;
                        Int64 DonorKey;

                        if (FindDonorByAccountNumber(PartnerByBankAccount.DefaultView, row.BranchCode, row.BankAccountNumber, out DonorShortName, out DonorKey))
                        {
                            tempRow.DonorKey = DonorKey;
                            tempRow.DonorShortName = DonorShortName;
                        }

#if disabled
                        // fuzzy search for the partner. only return if unique result
                        string sql =
                            "SELECT p_partner_key_n, p_partner_short_name_c FROM p_partner WHERE p_partner_short_name_c LIKE '{0}%' OR p_partner_short_name_c LIKE '{1}%'";
                        string[] names = row.AccountName.Split(new char[] { ' ' });

                        if (names.Length > 1)
                        {
                            string optionShortName1 = names[0] + ", " + names[1];
                            string optionShortName2 = names[1] + ", " + names[0];

                            DataTable partner = DBAccess.GDBAccessObj.SelectDT(String.Format(sql,
                                    optionShortName1,
                                    optionShortName2), "partner", Transaction);

                            if (partner.Rows.Count == 1)
                            {
                                tempRow.DonorKey = Convert.ToInt64(partner.Rows[0][0]);
                            }
                        }
#endif

                        MatchesToAddLater.Add(tempRow.MatchText, tempRow);

                        // do not modify tempRow.MatchAction, because that will not be stored in the database anyway, just costs time
                        row.MatchAction = tempRow.Action;
                    }
                }

                // for speed reasons, add the new rows after clearing the sort on the view
                TempDataset.AEpMatch.DefaultView.Sort = string.Empty;

                foreach (AEpMatchRow m in MatchesToAddLater.Values)
                {
                    TempDataset.AEpMatch.Rows.Add(m);
                }

                foreach (MatchDate date in NewDates)
                {
                    date.r.RecentMatch = date.d;
                }

                foreach (AEpMatchRow r in SetUnmatched)
                {
                    r.Action = MFinanceConstants.BANK_STMT_STATUS_UNMATCHED;
                }

                foreach (MatchDonor d in FindDonorKey)
                {
                    String DonorShortName;
                    Int64 DonorKey;

                    if (FindDonorByAccountNumber(PartnerByBankAccount.DefaultView, d.b, d.a, out DonorShortName, out DonorKey))
                    {
                        d.r.DonorKey = DonorKey;
                        d.r.DonorShortName = DonorShortName;
                    }
                }

                TProgressTracker.SetCurrentState(MyClientID,
                    Catalog.GetString("save matches"),
                    90.0m);

                TempDataset.ThrowAwayAfterSubmitChanges = true;
                // only store a_ep_transactions and a_ep_matches, but without additional typed fields (ie MatchAction)
                BankImportTDSAccess.SubmitChanges(TempDataset.GetChangesTyped(true));
            }
            catch (Exception e)
            {
                TLogging.Log(e.GetType().ToString() + " in BankImport, GetBankStatementTransactionsAndMatches; " + e.Message);
                TLogging.Log(e.StackTrace);
                DBAccess.GDBAccessObj.RollbackTransaction();
                throw;
            }

            // drop all matches that do not occur on this statement
            ResultDataset.AEpMatch.Clear();

            // reloading is faster than deleting all matches that are not needed
            string sqlLoadMatchesOfStatement =
                "SELECT DISTINCT m.* FROM PUB_a_ep_match m, PUB_a_ep_transaction t WHERE t.a_statement_key_i = ? AND m.a_ledger_number_i = ? AND m.a_match_text_c = t.a_match_text_c";

            OdbcParameter param = new OdbcParameter("statementkey", OdbcType.Int);
            param.Value = AStatementKey;
            OdbcParameter paramLedgerNumber = new OdbcParameter("ledgerNumber", OdbcType.Int);
            paramLedgerNumber.Value = ALedgerNumber;

            DBAccess.GDBAccessObj.SelectDT(ResultDataset.AEpMatch,
                sqlLoadMatchesOfStatement,
                null,
                new OdbcParameter[] { param, paramLedgerNumber }, -1, -1);

            // update the custom field for cost centre name for each match
            foreach (BankImportTDSAEpMatchRow row in ResultDataset.AEpMatch.Rows)
            {
                ACostCentreRow ccRow = (ACostCentreRow)ResultDataset.ACostCentre.Rows.Find(new object[] { row.LedgerNumber, row.CostCentreCode });

                if (ccRow != null)
                {
                    row.CostCentreName = ccRow.CostCentreName;
                }
            }

            // remove all rows that we do not need on the client side
            ResultDataset.AGiftDetail.Clear();
            ResultDataset.AMotivationDetail.Clear();
            ResultDataset.ACostCentre.Clear();

            ResultDataset.AcceptChanges();

            if (TLogging.DebugLevel > 0)
            {
                int CountMatched = 0;

                foreach (BankImportTDSAEpTransactionRow transaction in ResultDataset.AEpTransaction.Rows)
                {
                    if (!transaction.IsMatchActionNull() && (transaction.MatchAction != MFinanceConstants.BANK_STMT_STATUS_UNMATCHED))
                    {
                        CountMatched++;
                    }
                }

                TLogging.Log(
                    "Loading bank statement: matched: " + CountMatched.ToString() + " of " + ResultDataset.AEpTransaction.Rows.Count.ToString());
            }

            TProgressTracker.FinishJob(MyClientID);

            return ResultDataset;
        }

        /// <summary>
        /// commit matches into a_ep_match
        /// </summary>
        /// <param name="AMainDS"></param>
        /// <returns></returns>
        [RequireModulePermission("FINANCE-1")]
        public static bool CommitMatches(BankImportTDS AMainDS)
        {
            if (AMainDS == null)
            {
                return false;
            }

            bool NewTransaction;
            TDBTransaction Transaction = DBAccess.GetDBAccessObj((TDataBase)null).GetNewOrExistingTransaction(IsolationLevel.Serializable, out NewTransaction);

            try
            {
                for (int RowCount = 0; (RowCount != AMainDS.AEpMatch.Rows.Count); RowCount++)
                {
                    DataRow TheRow = AMainDS.AEpMatch.Rows[RowCount];
    
                    if (TheRow.RowState == DataRowState.Deleted)
                    {
                        string sql = "UPDATE " + AEpTransactionTable.GetTableDBName() +
                            " SET " + AEpTransactionTable.GetEpMatchKeyDBName() + " = NULL" +
                            " WHERE " + AEpTransactionTable.GetEpMatchKeyDBName() + " = " +
                            TheRow[AEpMatchTable.GetEpMatchKeyDBName(), DataRowVersion.Original];
                        DBAccess.GetDBAccessObj(Transaction).ExecuteNonQuery(sql, Transaction);
                    }
                }

                AMainDS.ThrowAwayAfterSubmitChanges = true;

                BankImportTDSAccess.SubmitChanges(AMainDS, Transaction.DataBaseObj);

                if (NewTransaction)
                {
                    DBAccess.GetDBAccessObj(Transaction.DataBaseObj).CommitTransaction();
                }

                return true;
            }
            catch (Exception e)
            {
                TLogging.Log("Bankimport, CommitMatches: " + e.ToString());

                if (NewTransaction)
                {
                    DBAccess.GetDBAccessObj(Transaction.DataBaseObj).RollbackTransaction();
                }

                return false;
            }
        }

        /// <summary>
        /// create a gift batch for the matched gifts, and return the gift batch number
        /// </summary>
        /// <returns>the gift batch number</returns>
        [RequireModulePermission("FINANCE-1")]
        public static Int32 CreateGiftBatch(
            Int32 ALedgerNumber,
            Int32 AStatementKey,
            Int32 AGiftBatchNumber,
            out TVerificationResultCollection AVerificationResult)
        {
            BankImportTDS MainDS = GetBankStatementTransactionsAndMatches(AStatementKey, ALedgerNumber);

            string MyClientID = DomainManager.GClientID.ToString();

            TProgressTracker.InitProgressTracker(MyClientID,
                Catalog.GetString("Creating gift batch"),
                MainDS.AEpTransaction.DefaultView.Count + 10);

            AVerificationResult = new TVerificationResultCollection();

            MainDS.AEpTransaction.DefaultView.RowFilter =
                String.Format("{0}={1}",
                    AEpTransactionTable.GetStatementKeyDBName(),
                    AStatementKey);
            MainDS.AEpStatement.DefaultView.RowFilter =
                String.Format("{0}={1}",
                    AEpStatementTable.GetStatementKeyDBName(),
                    AStatementKey);
            AEpStatementRow stmt = (AEpStatementRow)MainDS.AEpStatement.DefaultView[0].Row;

            // TODO: optional: use the preselected gift batch, AGiftBatchNumber

            Int32 DateEffectivePeriodNumber, DateEffectiveYearNumber;
            DateTime BatchDateEffective = stmt.Date;
            TDBTransaction Transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadCommitted);

            if (!TFinancialYear.GetLedgerDatePostingPeriod(ALedgerNumber, ref BatchDateEffective, out DateEffectiveYearNumber,
                    out DateEffectivePeriodNumber,
                    Transaction, true))
            {
                // just use the latest possible date
                string msg =
                    String.Format(Catalog.GetString("Date {0} is not in an open period of the ledger, using date {1} instead for the gift batch."),
                        stmt.Date.ToShortDateString(),
                        BatchDateEffective.ToShortDateString());
                AVerificationResult.Add(new TVerificationResult(Catalog.GetString("Creating Gift Batch"), msg, TResultSeverity.Resv_Info));
            }

            ACostCentreAccess.LoadViaALedger(MainDS, ALedgerNumber, Transaction);
            AMotivationDetailAccess.LoadViaALedger(MainDS, ALedgerNumber, Transaction);

            MainDS.AEpMatch.DefaultView.Sort =
                AEpMatchTable.GetActionDBName() + ", " +
                AEpMatchTable.GetMatchTextDBName();

            if (MainDS.AEpTransaction.DefaultView.Count == 0)
            {
                AVerificationResult.Add(new TVerificationResult(
                        Catalog.GetString("Creating Gift Batch"),
                        String.Format(Catalog.GetString("There are no transactions for statement #{0}."), AStatementKey),
                        TResultSeverity.Resv_Info));
                TProgressTracker.FinishJob(MyClientID);
                return -1;
            }

            foreach (DataRowView dv in MainDS.AEpTransaction.DefaultView)
            {
                AEpTransactionRow transactionRow = (AEpTransactionRow)dv.Row;

                DataRowView[] matches = MainDS.AEpMatch.DefaultView.FindRows(new object[] {
                        MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT,
                        transactionRow.MatchText
                    });

                if (matches.Length > 0)
                {
                    AEpMatchRow match = (AEpMatchRow)matches[0].Row;

                    if (match.IsDonorKeyNull() || (match.DonorKey == 0))
                    {
                        string msg =
                            String.Format(Catalog.GetString("Cannot create a gift for transaction {0} since there is no valid donor."),
                                transactionRow.Description);
                        AVerificationResult.Add(new TVerificationResult(Catalog.GetString("Creating Gift Batch"), msg, TResultSeverity.Resv_Critical));
                        DBAccess.GDBAccessObj.RollbackTransaction();
                        TProgressTracker.FinishJob(MyClientID);
                        return -1;
                    }
                }
            }

            string MatchedGiftReference = stmt.Filename;

            if (!stmt.IsBankAccountKeyNull())
            {
                string sqlGetBankSortCode =
                    "SELECT bank.p_branch_code_c " +
                    "FROM PUB_p_banking_details details, PUB_p_bank bank " +
                    "WHERE details.p_banking_details_key_i = ?" +
                    "AND details.p_bank_key_n = bank.p_partner_key_n";
                OdbcParameter param = new OdbcParameter("detailkey", OdbcType.Int);
                param.Value = stmt.BankAccountKey;

                PBankTable bankTable = new PBankTable();
                DBAccess.GDBAccessObj.SelectDT(bankTable, sqlGetBankSortCode, Transaction, new OdbcParameter[] { param }, 0, 0);

                MatchedGiftReference = bankTable[0].BranchCode + " " + stmt.Date.Day.ToString();
            }

            DBAccess.GDBAccessObj.RollbackTransaction();

            GiftBatchTDS GiftDS = TGiftTransactionWebConnector.CreateAGiftBatch(
                ALedgerNumber,
                BatchDateEffective,
                String.Format(Catalog.GetString("bank import for date {0}"), stmt.Date.ToShortDateString()));

            AGiftBatchRow giftbatchRow = GiftDS.AGiftBatch[0];
            giftbatchRow.BankAccountCode = stmt.BankAccountCode;

            decimal HashTotal = 0.0M;

            MainDS.AEpTransaction.DefaultView.Sort =
                AEpTransactionTable.GetNumberOnPaperStatementDBName();

            MainDS.AEpMatch.DefaultView.RowFilter = String.Empty;
            MainDS.AEpMatch.DefaultView.Sort =
                AEpMatchTable.GetActionDBName() + ", " +
                AEpMatchTable.GetMatchTextDBName();

            int counter = 5;

            foreach (DataRowView dv in MainDS.AEpTransaction.DefaultView)
            {
                TProgressTracker.SetCurrentState(MyClientID,
                    Catalog.GetString("Preparing the gifts"),
                    counter++);

                AEpTransactionRow transactionRow = (AEpTransactionRow)dv.Row;

                DataRowView[] matches = MainDS.AEpMatch.DefaultView.FindRows(new object[] {
                        MFinanceConstants.BANK_STMT_STATUS_MATCHED_GIFT,
                        transactionRow.MatchText
                    });

                if (matches.Length > 0)
                {
                    AEpMatchRow match = (AEpMatchRow)matches[0].Row;

                    AGiftRow gift = GiftDS.AGift.NewRowTyped();
                    gift.LedgerNumber = giftbatchRow.LedgerNumber;
                    gift.BatchNumber = giftbatchRow.BatchNumber;
                    gift.GiftTransactionNumber = giftbatchRow.LastGiftNumber + 1;
                    gift.DonorKey = match.DonorKey;
                    gift.DateEntered = transactionRow.DateEffective;
                    gift.Reference = MatchedGiftReference;
                    GiftDS.AGift.Rows.Add(gift);
                    giftbatchRow.LastGiftNumber++;

                    foreach (DataRowView r in matches)
                    {
                        match = (AEpMatchRow)r.Row;

                        AGiftDetailRow detail = GiftDS.AGiftDetail.NewRowTyped();
                        detail.LedgerNumber = gift.LedgerNumber;
                        detail.BatchNumber = gift.BatchNumber;
                        detail.GiftTransactionNumber = gift.GiftTransactionNumber;
                        detail.DetailNumber = gift.LastDetailNumber + 1;
                        gift.LastDetailNumber++;

                        detail.GiftTransactionAmount = match.GiftTransactionAmount;
                        detail.GiftAmount = match.GiftTransactionAmount;
                        HashTotal += match.GiftTransactionAmount;
                        detail.MotivationGroupCode = match.MotivationGroupCode;
                        detail.MotivationDetailCode = match.MotivationDetailCode;

                        // do not use the description in comment one, because that could show up on the gift receipt???
                        // detail.GiftCommentOne = transactionRow.Description;

                        detail.CommentOneType = MFinanceConstants.GIFT_COMMENT_TYPE_BOTH;
                        detail.CostCentreCode = match.CostCentreCode;
                        detail.RecipientKey = match.RecipientKey;
                        detail.RecipientLedgerNumber = match.RecipientLedgerNumber;

                        AMotivationDetailRow motivation = (AMotivationDetailRow)MainDS.AMotivationDetail.Rows.Find(
                            new object[] { ALedgerNumber, detail.MotivationGroupCode, detail.MotivationDetailCode });

                        if (motivation == null)
                        {
                            AVerificationResult.Add(new TVerificationResult(
                                    String.Format(Catalog.GetString("creating gift for match {0}"), transactionRow.Description),
                                    String.Format(Catalog.GetString("Cannot find motivation group '{0}' and motivation detail '{1}'"),
                                        detail.MotivationGroupCode, detail.MotivationDetailCode),
                                    TResultSeverity.Resv_Critical));
                        }

                        // try to retrieve the current costcentre for this recipient
                        if (detail.RecipientKey != 0)
                        {
                            detail.RecipientLedgerNumber = TGiftTransactionWebConnector.GetRecipientFundNumber(detail.RecipientKey);

                            detail.CostCentreCode = TGiftTransactionWebConnector.IdentifyPartnerCostCentre(detail.LedgerNumber,
                                detail.RecipientLedgerNumber);
                        }
                        else
                        {
                            if (motivation != null)
                            {
                                detail.CostCentreCode = motivation.CostCentreCode;
                            }
                        }

                        // check for active cost centre
                        ACostCentreRow costcentre = (ACostCentreRow)MainDS.ACostCentre.Rows.Find(new object[] { ALedgerNumber, detail.CostCentreCode });

                        if ((costcentre == null) || !costcentre.CostCentreActiveFlag)
                        {
                            AVerificationResult.Add(new TVerificationResult(
                                    String.Format(Catalog.GetString("creating gift for match {0}"), transactionRow.Description),
                                    Catalog.GetString("Invalid or inactive cost centre"),
                                    TResultSeverity.Resv_Critical));
                        }

                        GiftDS.AGiftDetail.Rows.Add(detail);
                    }
                }
            }

            TProgressTracker.SetCurrentState(MyClientID,
                Catalog.GetString("Submit to database"),
                counter++);

            if (AVerificationResult.HasCriticalErrors)
            {
                TProgressTracker.FinishJob(MyClientID);
                return -1;
            }

            giftbatchRow.HashTotal = HashTotal;
            giftbatchRow.BatchTotal = HashTotal;

            // do not overwrite the parameter, because there might be the hint for a different gift batch date
            TVerificationResultCollection VerificationResultSubmitChanges;

            TSubmitChangesResult result = TGiftTransactionWebConnector.SaveGiftBatchTDS(ref GiftDS,
                out VerificationResultSubmitChanges);

            TProgressTracker.FinishJob(MyClientID);

            if (result == TSubmitChangesResult.scrOK)
            {
                return giftbatchRow.BatchNumber;
            }
            else
            {
                TLogging.Log("Bankimport: Problem with creating gift batch:");
                TLogging.Log(VerificationResultSubmitChanges.BuildVerificationResultString());
            }

            return -1;
        }

        /// <summary>
        /// create a GL batch for the matched GL Transactions, and return the GL batch number
        /// </summary>
        /// <returns>the GL batch number</returns>
        [RequireModulePermission("FINANCE-1")]
        public static Int32 CreateGLBatch(BankImportTDS AMainDS,
            Int32 ALedgerNumber,
            Int32 AStatementKey,
            Int32 AGLBatchNumber,
            out TVerificationResultCollection AVerificationResult)
        {
            AMainDS.AEpTransaction.DefaultView.RowFilter =
                String.Format("{0}={1}",
                    AEpTransactionTable.GetStatementKeyDBName(),
                    AStatementKey);
            AMainDS.AEpStatement.DefaultView.RowFilter =
                String.Format("{0}={1}",
                    AEpStatementTable.GetStatementKeyDBName(),
                    AStatementKey);
            AEpStatementRow stmt = (AEpStatementRow)AMainDS.AEpStatement.DefaultView[0].Row;

            AVerificationResult = null;

            Int32 DateEffectivePeriodNumber, DateEffectiveYearNumber;
            TDBTransaction Transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadCommitted);

            if (!TFinancialYear.IsValidPostingPeriod(ALedgerNumber, stmt.Date, out DateEffectivePeriodNumber, out DateEffectiveYearNumber,
                    Transaction))
            {
                string msg = String.Format(Catalog.GetString("Cannot create a GL batch for date {0} since it is not in an open period of the ledger."),
                    stmt.Date.ToShortDateString());
                TLogging.Log(msg);
                AVerificationResult = new TVerificationResultCollection();
                AVerificationResult.Add(new TVerificationResult(Catalog.GetString("Creating GL Batch"), msg, TResultSeverity.Resv_Critical));

                DBAccess.GDBAccessObj.RollbackTransaction();
                return -1;
            }

            Int32 BatchYear, BatchPeriod;

            // if DateEffective is outside the range of open periods, use the most fitting date
            DateTime DateEffective = stmt.Date;
            TFinancialYear.GetLedgerDatePostingPeriod(ALedgerNumber, ref DateEffective, out BatchYear, out BatchPeriod, Transaction, true);

            ALedgerTable LedgerTable = ALedgerAccess.LoadByPrimaryKey(ALedgerNumber, Transaction);

            DBAccess.GDBAccessObj.RollbackTransaction();

            GLBatchTDS GLDS = TGLTransactionWebConnector.CreateABatch(ALedgerNumber);

            ABatchRow glbatchRow = GLDS.ABatch[0];
            glbatchRow.BatchPeriod = BatchPeriod;
            glbatchRow.DateEffective = DateEffective;
            glbatchRow.BatchDescription = String.Format(Catalog.GetString("bank import for date {0}"), stmt.Date.ToShortDateString());

            decimal HashTotal = 0.0M;
            decimal DebitTotal = 0.0M;
            decimal CreditTotal = 0.0M;

            // TODO: support several journals
            // TODO: support several currencies, support other currencies than the base currency
            AJournalRow gljournalRow = GLDS.AJournal.NewRowTyped();
            gljournalRow.LedgerNumber = glbatchRow.LedgerNumber;
            gljournalRow.BatchNumber = glbatchRow.BatchNumber;
            gljournalRow.JournalNumber = glbatchRow.LastJournal + 1;
            gljournalRow.TransactionCurrency = LedgerTable[0].BaseCurrency;
            glbatchRow.LastJournal++;
            gljournalRow.JournalPeriod = glbatchRow.BatchPeriod;
            gljournalRow.DateEffective = glbatchRow.DateEffective;
            gljournalRow.JournalDescription = glbatchRow.BatchDescription;
            gljournalRow.SubSystemCode = CommonAccountingSubSystemsEnum.GL.ToString();
            gljournalRow.TransactionTypeCode = CommonAccountingTransactionTypesEnum.STD.ToString();
            gljournalRow.ExchangeRateToBase = 1.0m;
            GLDS.AJournal.Rows.Add(gljournalRow);

            ATransactionRow trans;

            foreach (DataRowView dv in AMainDS.AEpTransaction.DefaultView)
            {
                AEpTransactionRow transactionRow = (AEpTransactionRow)dv.Row;

                DataView v = AMainDS.AEpMatch.DefaultView;
                v.RowFilter = AEpMatchTable.GetActionDBName() + " = '" + MFinanceConstants.BANK_STMT_STATUS_MATCHED_GL + "' and " +
                              AEpMatchTable.GetMatchTextDBName() + " = '" + transactionRow.MatchText + "'";

                if (v.Count > 0)
                {
                    AEpMatchRow match = (AEpMatchRow)v[0].Row;
                    trans = GLDS.ATransaction.NewRowTyped();
                    trans.LedgerNumber = glbatchRow.LedgerNumber;
                    trans.BatchNumber = glbatchRow.BatchNumber;
                    trans.JournalNumber = gljournalRow.JournalNumber;
                    trans.TransactionNumber = gljournalRow.LastTransactionNumber + 1;
                    trans.AccountCode = match.AccountCode;
                    trans.CostCentreCode = match.CostCentreCode;
                    trans.Reference = match.Reference;
                    trans.Narrative = match.Narrative;
                    trans.TransactionDate = transactionRow.DateEffective;

                    if (transactionRow.TransactionAmount < 0)
                    {
                        trans.AmountInBaseCurrency = -1 * transactionRow.TransactionAmount;
                        trans.TransactionAmount = -1 * transactionRow.TransactionAmount;
                        trans.DebitCreditIndicator = true;
                        DebitTotal += trans.AmountInBaseCurrency;
                    }
                    else
                    {
                        trans.AmountInBaseCurrency = transactionRow.TransactionAmount;
                        trans.TransactionAmount = transactionRow.TransactionAmount;
                        trans.DebitCreditIndicator = false;
                        CreditTotal += trans.AmountInBaseCurrency;
                    }

                    GLDS.ATransaction.Rows.Add(trans);
                    gljournalRow.LastTransactionNumber++;
                }
            }

            // add one transaction for the bank account as well
            trans = GLDS.ATransaction.NewRowTyped();
            trans.LedgerNumber = glbatchRow.LedgerNumber;
            trans.BatchNumber = glbatchRow.BatchNumber;
            trans.JournalNumber = gljournalRow.JournalNumber;
            trans.TransactionNumber = gljournalRow.LastTransactionNumber + 1;
            trans.AccountCode = stmt.BankAccountCode;
            trans.CostCentreCode = TLedgerInfo.GetStandardCostCentre(ALedgerNumber);
            trans.Reference = String.Empty;
            trans.Narrative = gljournalRow.JournalDescription;
            trans.TransactionDate = glbatchRow.DateEffective;

            if (CreditTotal > DebitTotal)
            {
                trans.AmountInBaseCurrency = CreditTotal - DebitTotal;
                trans.TransactionAmount = CreditTotal - DebitTotal;
                trans.DebitCreditIndicator = true;
                DebitTotal += (CreditTotal - DebitTotal);
            }
            else
            {
                trans.AmountInBaseCurrency = DebitTotal - CreditTotal;
                trans.TransactionAmount = DebitTotal - CreditTotal;
                trans.DebitCreditIndicator = false;
                CreditTotal += (DebitTotal - CreditTotal);
            }

            GLDS.ATransaction.Rows.Add(trans);
            gljournalRow.LastTransactionNumber++;

            gljournalRow.JournalDebitTotal = DebitTotal;
            gljournalRow.JournalCreditTotal = CreditTotal;
            glbatchRow.BatchDebitTotal = DebitTotal;
            glbatchRow.BatchCreditTotal = CreditTotal;
            glbatchRow.BatchControlTotal = HashTotal;

            TVerificationResultCollection VerificationResult;

            TSubmitChangesResult result = TGLTransactionWebConnector.SaveGLBatchTDS(ref GLDS,
                out VerificationResult);

            if (result == TSubmitChangesResult.scrOK)
            {
                return glbatchRow.BatchNumber;
            }

            TLogging.Log("Problems storing GL Batch");
            return -1;
        }
    }
}