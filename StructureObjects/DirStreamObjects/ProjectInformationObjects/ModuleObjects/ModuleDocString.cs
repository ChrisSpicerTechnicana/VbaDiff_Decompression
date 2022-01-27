using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.4
    /// </summary>
    internal class ModuleDocString
    {
        internal void ParseStream(byte[] stream, ref int position)
        { 
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x001C) { throw new ParseException("Failed to parse ID in ModuleDocString."); }

            // Size of DocString
            uint sizeOfDocString = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in doc string at the moment. Just move on.
            position += (int)sizeOfDocString;

            // Reserved
            uint reserved = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved != 0x0048) { throw new ParseException("Failed to parsed reserved in ModuleDocString."); }

            // Size of DocString in Unicode
            uint sizeOfDocStringUnicode = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in Doc string unicode just yet.
            position += (int)sizeOfDocStringUnicode;

        }
    }
}
