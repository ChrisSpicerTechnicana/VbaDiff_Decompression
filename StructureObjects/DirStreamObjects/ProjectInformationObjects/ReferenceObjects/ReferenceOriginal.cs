using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// 2.3.4.2.2.4
    /// </summary>
    class ReferenceOriginal
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0033) { throw new ParseException("Failed to parse the ReferenceOriginal ID."); }

            // SizeOfLibidOriginal
            uint sizeOfLibidOriginal = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in the LibId original right now
            position += (int)sizeOfLibidOriginal;
        }
    }
}
