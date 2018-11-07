using System;

namespace flookup_exception
{
    /// <summary>
    /// Exception is thrown if the specified inbox cannot be found
    /// </summary>
    public class InvalidFolderException : Exception
    {
        public InvalidFolderException(string message) : base(message)
        {

        }
    }
}
