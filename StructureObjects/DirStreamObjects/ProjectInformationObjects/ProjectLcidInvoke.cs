using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    class ProjectLcidInvoke
    {
        internal bool ParseStream(byte[] stream, ref int position)
        {
            // Read the Id
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0014) { throw new ParseException(); }

            // Read the Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000004) { throw new ParseException(); }

            // Read the LcidInvoke
            uint lcidInvoke = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (lcidInvoke != 0x00000409) { throw new ParseException(); }

            return true;
        }
    }
}
