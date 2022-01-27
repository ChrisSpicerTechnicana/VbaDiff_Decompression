using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    class ProjectHelpFilePath
    {
        internal void ParseStream(byte[] stream, ref int position)
        { 
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0006) { throw new ParseException("Failed to parse ID for ProjectHelpFilePath"); }

            // Size of help file 1.
            uint sizeOfHelpFile1 = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (sizeOfHelpFile1 > 260) { throw new ParseException("sizeOfHelpFile1 in ProjectHelpFilePath is too big."); }

            // Ignore the Helpfile1
            position +=(int) sizeOfHelpFile1;

            // Reserved
            uint reserved = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved != 0x003d) { throw new ParseException("Reserved field in ProjectHelpFilePath is wrong."); }

            // Size of help file 2.
            uint sizeOfHelpFile2 = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (sizeOfHelpFile2 != sizeOfHelpFile1) { throw new ParseException("sizeOfHelpFile2 does not match sizeOfHelpFile1 in ProjectHelpFilePath."); }

            // Ignore the Helpfile2
             position += (int)sizeOfHelpFile2;

        }
    }
}
