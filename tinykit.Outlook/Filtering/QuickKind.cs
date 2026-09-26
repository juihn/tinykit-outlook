namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>Quick-filter attribute; the value is matched as "contains".</summary>
    internal enum QuickKind
    {
        /// <summary>The From (sender) e-mail address.</summary>
        From,

        /// <summary>The nameRelated column.</summary>
        Name,

        /// <summary>The subject.</summary>
        Subject,

        /// <summary>The domainRelated column.</summary>
        Domain,
    }
}
