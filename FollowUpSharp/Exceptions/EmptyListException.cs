using System;
using System.Runtime.Serialization;

namespace FollowUpSharp.Exceptions
{
    class EmptyListException : Exception
    {
        public EmptyListException() : base() { }
        public EmptyListException(string message) : base(message) { }
        public EmptyListException(string message, Exception inner) : base(message, inner) { }

        protected EmptyListException(SerializationInfo info, StreamingContext context) { }
    }
}
