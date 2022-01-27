using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.10
    /// </summary>
    class ProjectVersion
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0009) { throw new ParseException("Parse failed for id in ProjectVersion."); }

            // Reserved 
            uint reserved = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (reserved != 0x00000004) { throw new ParseException("Parse failed for reserved in ProjectVersion."); }

            // Major Version Number
            uint versionMajor = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Minor Version Number
            uint versionMinor = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;
        }
    }
}
