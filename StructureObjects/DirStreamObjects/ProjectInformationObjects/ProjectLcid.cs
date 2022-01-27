using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.2
    /// </summary>
    class ProjectLcid
    {
        internal bool ParseStream(byte[] stream, ref int position)
        {
            // Read the ID.
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0002) { throw new ParseException(); }

            // Read the Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000004) { throw new ParseException(); }
            
            // Read the Lcid
            uint lcid = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (lcid != 0x00000409) { throw new ParseException(); }

            return true;
        }
    }
}
