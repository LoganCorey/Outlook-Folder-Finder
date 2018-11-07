using System;

namespace flookup_exception
{
    /// <summary>
    /// Exception is thrown if the specified inbox cannot be found
    /// </summary>
    public class InvalidInboxException : Exception
    {
        public InvalidInboxException(string message) : base(message)
        {

        }
    }
}
