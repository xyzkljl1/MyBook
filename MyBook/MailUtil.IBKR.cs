using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MimeKit;

namespace MyBook
{
    // IBKR 邮件报表发现、附件筛选与账户匹配逻辑。
    partial class MailUtil
    {
        private const StatementImportProvider IBKRProvider = StatementImportProvider.IBKRReportMail;
        private const string IBKRReportSender = "donotreply@interactivebrokers.com";
        private const string IBKRDailyReportType = "DailyMyBook";
        private const int IBKRMissingReportLimitDays = 14;

        private async Task FetchIBKRReports()
        {
            var date = GetNextDailyStatementDate(IBKRProvider);
            var missingDays = 0;
            while (date <= DateTime.Today)
            {
                var imported = await FetchIBKRReport(date);
                if (imported)
                {
                    missingDays = 0;
                }
                else
                {
                    missingDays++;
                    if (missingDays >= IBKRMissingReportLimitDays)
                        throw new InvalidOperationException($"Missing IBKR reports for {IBKRMissingReportLimitDays} consecutive days ending {date:yyyy-MM-dd}");
                }

                date = date.AddDays(1);
            }
        }

        private async Task<bool> FetchIBKRReport(DateTime date)
        {
            var subjectDate = date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            var expectedSubject = $"{subjectDate}的自定义活动报表";
            if (database.IsStatementImported(IBKRProvider, date))
                return true;

            var message = await SearchBill(
                IBKRReportSender,
                expectedSubject,
                date,
                message => String.Equals(message.Subject?.Trim(), expectedSubject, StringComparison.Ordinal)
                    && HasIBKRReportAttachment(message, IBKRDailyReportType, date));

            if (message is null)
            {
                Console.WriteLine($"Find no IBKR {IBKRDailyReportType} report {subjectDate}");
                return false;
            }

            var reportAttachments = ReadIBKRReportAttachments(message, IBKRDailyReportType, date);
            if (reportAttachments.Count == 0)
            {
                Console.WriteLine($"parse IBKR report mail fail: no {IBKRDailyReportType} attachment, subject={message.Subject}");
                throw new MailParseException($"Parse IBKR Report Fail, Missing {IBKRDailyReportType} Attachment: {message.Subject}");
            }

            var reportDate = reportAttachments[0].ReportDate;
            if (reportDate.Date != date.Date)
                throw new MailParseException($"Parse IBKR Report Fail, Date Mismatch: expected {date:yyyy-MM-dd}, got {reportDate:yyyy-MM-dd}");

            var bodyText = GetMessageText(message);
            var reportAccountId = ParseIBKRAccountId(bodyText);
            if (String.IsNullOrWhiteSpace(reportAccountId))
            {
                Console.WriteLine($"parse IBKR report mail fail: missing account id, subject={message.Subject}");
                throw new MailParseException($"Parse IBKR Report Fail, Missing Account: {message.Subject}");
            }

            var account = GetIBKRAccount(reportAccountId);

            Console.WriteLine($"Find IBKR {IBKRDailyReportType} report: {reportDate:yyyy-MM-dd} {reportAccountId} -> {account.name}, attachments={reportAttachments.Count}");
            foreach (var attachment in reportAttachments)
                Console.WriteLine($"Load IBKR report in memory: {attachment.FileName}, id={attachment.ReportId}, bytes={attachment.Content.Length}");

            // TODO: parse reportAttachments and save report content.
            database.SaveStatementImportOnce(IBKRProvider, reportDate);
            return true;
        }

        private Account GetIBKRAccount(string reportAccountId)
        {
            var normalizedReportAccountId = reportAccountId.Trim().ToUpperInvariant();
            var visibleDigits = Regex.Match(normalizedReportAccountId, @"\d+$").Value;
            if (String.IsNullOrWhiteSpace(visibleDigits))
                throw new InvalidOperationException($"Invalid IBKR account id: {reportAccountId}");

            var candidates = database.GetAllAccounts()
                .Where(account => account.name.StartsWith("IBKR_", StringComparison.OrdinalIgnoreCase))
                .Where(account =>
                {
                    var accountId = account.name["IBKR_".Length..].Trim().ToUpperInvariant();
                    if (normalizedReportAccountId.Contains('*'))
                        return accountId.EndsWith(visibleDigits, StringComparison.OrdinalIgnoreCase);
                    return accountId.Equals(normalizedReportAccountId, StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            if (candidates.Count == 1)
                return candidates[0];

            var usdCandidates = candidates.Where(account => account._v_t == CurrencyType.USD).ToList();
            if (usdCandidates.Count == 1)
                return usdCandidates[0];

            if (candidates.Count > 1)
            {
                throw new InvalidOperationException(
                    $"Find multiple IBKR accounts for {reportAccountId}: {String.Join(",", candidates.Select(account => account.name))}");
            }

            throw new InvalidOperationException($"Account not found: IBKR/{reportAccountId}");
        }

        private static string? ParseIBKRAccountId(string text)
        {
            var normalizedText = NormalizeMailText(text);
            var labeledMatch = Regex.Match(
                normalizedText,
                @"(?:Account(?:\s*ID|\s*Number)?|账号|账户)\s*[:：]?\s*(?<account>U(?:\*+|\d*)\d{4,})",
                RegexOptions.IgnoreCase);
            if (labeledMatch.Success)
                return labeledMatch.Groups["account"].Value.ToUpperInvariant();

            var match = Regex.Match(normalizedText, @"\b(?<account>U(?:\*{2,}\d{4,}|\d{5,}))\b", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups["account"].Value.ToUpperInvariant() : null;
        }

        private static List<InMemoryIBKRReportAttachment> ReadIBKRReportAttachments(
            MimeMessage message,
            string reportType,
            DateTime reportDate)
        {
            var attachments = new List<InMemoryIBKRReportAttachment>();
            foreach (var attachment in message.Attachments)
            {
                if (attachment is not MimePart mimePart)
                    continue;

                var fileName = GetAttachmentFileName(attachment);
                if (!TryParseIBKRReportAttachmentName(fileName, out var attachmentInfo)
                    || !String.Equals(attachmentInfo.ReportType, reportType, StringComparison.OrdinalIgnoreCase)
                    || attachmentInfo.ReportDate.Date != reportDate.Date)
                    continue;

                using var memory = new MemoryStream();
                mimePart.Content.DecodeTo(memory);
                attachments.Add(new InMemoryIBKRReportAttachment(
                    fileName,
                    attachmentInfo.ReportType,
                    attachmentInfo.ReportId,
                    attachmentInfo.ReportDate,
                    memory.ToArray()));
            }

            return attachments;
        }

        private static bool HasIBKRReportAttachment(MimeMessage message, string reportType, DateTime reportDate)
        {
            return message.Attachments.Any(attachment =>
            {
                var fileName = GetAttachmentFileName(attachment);
                return TryParseIBKRReportAttachmentName(fileName, out var attachmentInfo)
                    && String.Equals(attachmentInfo.ReportType, reportType, StringComparison.OrdinalIgnoreCase)
                    && attachmentInfo.ReportDate.Date == reportDate.Date;
            });
        }

        private static bool TryParseIBKRReportAttachmentName(string fileName, out IBKRReportAttachmentInfo attachmentInfo)
        {
            attachmentInfo = new IBKRReportAttachmentInfo("", "", default);
            var name = Path.GetFileName(fileName.Trim());
            var match = Regex.Match(name, @"^(?<type>[^.]+)\.(?<id>[^.]+)\.(?<date>\d{8})(?:\..*)?$");
            if (!match.Success)
                return false;

            if (!DateTime.TryParseExact(
                match.Groups["date"].Value,
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var reportDate))
            {
                return false;
            }

            attachmentInfo = new IBKRReportAttachmentInfo(
                match.Groups["type"].Value,
                match.Groups["id"].Value,
                reportDate);
            return true;
        }

        private static string GetAttachmentFileName(MimeEntity attachment)
        {
            return attachment.ContentDisposition?.FileName
                ?? attachment.ContentType.Name
                ?? "";
        }

        private sealed record IBKRReportAttachmentInfo(string ReportType, string ReportId, DateTime ReportDate);

        private sealed record InMemoryIBKRReportAttachment(
            string FileName,
            string ReportType,
            string ReportId,
            DateTime ReportDate,
            byte[] Content);
    }
}
