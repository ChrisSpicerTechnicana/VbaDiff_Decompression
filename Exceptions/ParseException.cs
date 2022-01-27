using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VbaDiff.Decompression.Exceptions
{
    class ParseException : Exception
    {
        internal ParseException()
        { }

        internal ParseException(string message, Exception innerException)
            :base(message, innerException)
        {
        }

        internal ParseException(string message)
            :base(message)
        {         
        }
    }
}
