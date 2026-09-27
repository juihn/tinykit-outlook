namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>Quick-filter attribute; the value is matched as "contains". Each has its own recent-value history.</summary>
    internal enum QuickKind
    {
        // ----- Mail -----

        /// <summary>The From (sender) e-mail address.</summary>
        From,

        /// <summary>The nameRelated column.</summary>
        Name,

        /// <summary>The subject.</summary>
        Subject,

        /// <summary>The domainRelated column.</summary>
        Domain,

        // ----- Contacts -----

        /// <summary>The contact's File As.</summary>
        FileAs,

        /// <summary>Any of the contact's three e-mail addresses.</summary>
        Email,

        /// <summary>The contact's company.</summary>
        Company,

        /// <summary>The contact's department.</summary>
        Department,
    }
}
