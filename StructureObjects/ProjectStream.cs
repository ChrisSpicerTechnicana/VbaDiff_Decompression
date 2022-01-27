using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Technicana.Infrastructure.Utilities.Strings;

namespace VbaDiff.Decompression.StructureObjects
{
    internal class ProjectStream
    {
        byte[] WriteProjectStream(string projectText, int codePage)
        {
            byte[] projectStream = StringUtilities.EncodeString(projectText, codePage);

            return projectStream;
        }
    }
}
