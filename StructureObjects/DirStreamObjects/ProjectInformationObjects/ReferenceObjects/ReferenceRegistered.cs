using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// 2.3.4.2.2.5
    /// </summary>
    class ReferenceRegistered
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x000D) { throw new ParseException("Failed to parseID in ReferenceRegistered."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // You have to check size laters.

            // SizeOfLibId
            uint sizeOfLibid = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in Libid right now.
            position += (int) sizeOfLibid;

            // Reserved 1
            uint reserved1 = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (reserved1 != 0x00000000) { throw new ParseException("Failed to parse Reserved1 field in ReferenceRegistered."); }

            // Reserved 2
            uint reserved2 = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved2 != 0x00000000) { throw new ParseException("Failed to parse Reserved1 field in ReferenceRegistered."); }


            // Now check the size
            if(size != (4 + sizeOfLibid + 4+ 2)) { throw new ParseException("The size doesn't match up in ReferenceRegistereed.");}
        }
    }
}
