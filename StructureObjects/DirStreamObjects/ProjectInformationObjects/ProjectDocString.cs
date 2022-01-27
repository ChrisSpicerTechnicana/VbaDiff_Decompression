using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.6
    /// </summary>
    class ProjectDocString
    {

        internal void ParseStream(byte[] stream, ref int position)
        {
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0005) { throw new ParseException("Failed to parse ProjectDocString Id."); }

            //Size of DocString
            uint sizeOfDocString = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (sizeOfDocString > 2000) { throw new ParseException(String.Format("DocString in ProjectDocStringRecord is too long. Length is {0}, should be 2000 or less.", sizeOfDocString)); }

            // I'm not currently interested in the DocString,and I can't run any checks on it, so just move on.
            position += (int)sizeOfDocString;

            // Reserved.
            uint reserved = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved != 0x0040) { throw new ParseException("Failed to parse projectDocString reserved field."); }

            // SizeOfDocStringUnicode
            uint sizeOfDocStringUnicode = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if ((sizeOfDocStringUnicode % 2) != 0) { throw new ParseException("Failed to parse ProjectDocstring sizeOfDocStringUnicode."); }

            // Again, I'm not that interested in the DocStringUnicode just yet, so I'll ignore it for now.
            position += (int)sizeOfDocStringUnicode;            
        }
    }
}
