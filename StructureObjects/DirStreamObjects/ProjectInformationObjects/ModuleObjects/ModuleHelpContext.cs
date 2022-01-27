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
    internal class ModuleHelpContext
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x001E) { throw new ParseException("Failed to parse Id in ModuleHelpContext."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000004) { throw new ParseException("Failed to parse size in ModuleHelpContext."); }

            // Help Context - of no interest right now.
            position += 4;
        }
    }
}
