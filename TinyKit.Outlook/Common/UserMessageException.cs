using System;

namespace TinyKit.OutlookAddin.Common
{
    /// <summary>An expected condition shown to the user as a plain message (not an error report).</summary>
    internal sealed class UserMessageException : Exception
    {
        public UserMessageException(string message) : base(message)
        {
        }
    }
}
