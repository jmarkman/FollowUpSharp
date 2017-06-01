using System;
using System.Runtime.Serialization;

namespace FollowUpSharp.Exceptions
{
    class EmptyQueryException : Exception
    {
        public EmptyQueryException() : base("There were no items to be returned from the query!") { }
        public EmptyQueryException(string message) : base(message) { }
        public EmptyQueryException(string message, Exception inner) : base(message, inner) { }

        protected EmptyQueryException(SerializationInfo info, StreamingContext context) { }
    }
}
