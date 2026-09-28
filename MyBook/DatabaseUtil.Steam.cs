using System.Text.RegularExpressions;

namespace MyBook;

partial class DatabaseUtil
{
    internal List<WebUtil.SteamBankPurchase> GetSteamBankPurchases(int steamAccountId)
    {
        var records = db.Queryable<Record>().Where(r => r.Source.Contains("code=SteamOrder-") || (r._account_Id != steamAccountId && r.v < 0
            && !r.Fake && !r.isInternal && !r.isRefundMatched && r.matchedRecordId == null && r.backup == null
            && (r.Source.Contains("steam") || r.Source.Contains("valve")
                || r.DestAccount.Contains("steam") || r.DestAccount.Contains("valve"))))
            .ToList().Where(r => WebUtil.HasSteamOrderCode(r.Source) || WebUtil.IsSteamBankPurchase(r)).ToList();
        if (records.Count == 0) return [];
        var accounts = db.Queryable<Account>().Select(a => new { a.Id, a._primaryAccount_Id }).ToList().ToDictionary(a => a.Id);
        var cards = db.Queryable<AccountInternalId>().Where(i => i._account_Id != null).ToList()
            .ToLookup(i => accounts[i._account_Id!.Value]._primaryAccount_Id ?? i._account_Id.Value,
                i => Regex.Replace(i.cardNo, @"[\s-]", ""));
        return records.Select(r => new WebUtil.SteamBankPurchase(r, cards[r._account_Id].ToArray())).ToList();
    }

    // Called inside the Steam import transaction, so matching uses the current bank records.
    internal int AppendSteamPurchaseSupplements(int steamAccountId, IReadOnlyList<WebUtil.SteamExternalPurchase> purchases)
    {
        var supplements = WebUtil.BuildSteamPurchaseSupplements(purchases, GetSteamBankPurchases(steamAccountId));
        AppendRecordSourceSupplements(supplements);
        return supplements.Count;
    }
}
