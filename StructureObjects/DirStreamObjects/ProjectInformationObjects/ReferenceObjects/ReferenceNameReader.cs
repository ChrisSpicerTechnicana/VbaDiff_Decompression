using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// 2.3.4.2.2.2
    /// </summary>
    internal static class ReferenceNameReader
    {       
        internal static RawReferenceName ParseStream(byte[] stream, ref int position)
        {
            // We'll populate this data object as we parse the stream.
            RawReferenceName rawReferenceName = new RawReferenceName();

            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0016) { throw new ParseException("Failed to parse ID in Reference Name."); }

            // Size of Name
            uint sizeOfName = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Name
            rawReferenceName.NameBytes = stream.SubArray(position, (int)sizeOfName);
            position += (int)sizeOfName;

            // Reserved
            uint reserved = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved != 0x003E) { throw new ParseException("Failed to parse Reserved field from ReferenceName."); }

            // Size of Name Unicode
            uint sizeOfNameUnicode = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in NameUnicode just yet.
            position += (int)sizeOfNameUnicode;

            return rawReferenceName;
        }
    }
}
