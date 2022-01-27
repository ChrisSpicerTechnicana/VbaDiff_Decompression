using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.8
    /// </summary>
    internal class ModuleType
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (!((id == 0x0021) || (id == 0x0022))) { throw new ParseException("Failed to parse Id in ModuleType."); }
 
            // TODO: Examine whether this module type is actually useful.

            uint reserved = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (reserved != 0x00000000) { throw new ParseException("Failed to parse reserved in ModuleType"); }

        }
    }
}
