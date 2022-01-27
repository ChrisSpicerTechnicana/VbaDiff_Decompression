using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.9
    /// </summary>
    class ProjectLibFlagsRecord
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            // Id
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0008) { throw new ParseException("Failed to parse the ID of ProjectLibFlagsRecord."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000004) { throw new ParseException("Failed to parse the Size of ProjectLibFlagsRecord."); }

            // ProjectLibFlags
            uint projectLibFlags = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (projectLibFlags != 0x00000000) { throw new ParseException("Failed the parse the ProjectLibs flag of ProjectLibFlagsRecord."); }
        }

    }
}
