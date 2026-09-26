namespace MyBook;

internal static class ImportSchedule
{
    internal const string SuccessfulQueryKey = "scheduled-query";

    // When supplied, the callback advances query progress even after an empty successful query.
    internal static async Task RunAsync(string name, int intervalDays, int missingAfterDays,
        Func<DateTime?> readProgress, Func<DateTime, Task> fetch, Action<DateTime>? markSuccessfulQuery = null,
        DateTime? today = null)
    {
        if (intervalDays <= 0 || missingAfterDays < 0
            || missingAfterDays > 0 && missingAfterDays <= intervalDays)
            throw new ArgumentOutOfRangeException(nameof(intervalDays));
        var date = (today ?? DateTime.Today).Date;
        var previous = readProgress()?.Date
            ?? throw new InvalidOperationException($"Missing {name} import checkpoint.");
        var elapsedDays = (date - previous).Days;
        if (elapsedDays < intervalDays)
            return;

        await fetch(previous).ConfigureAwait(false);
        if (missingAfterDays > 0 && elapsedDays > missingAfterDays
            && readProgress()?.Date <= previous)
            throw new InvalidOperationException(
                $"{name}: no new report imported for {elapsedDays} days (limit {missingAfterDays}); "
                + $"last source date {previous:yyyy-MM-dd}, checked through {date:yyyy-MM-dd}.");

        // Failed queries never reach this point; successful empty queries do.
        markSuccessfulQuery?.Invoke(date);
    }
}
