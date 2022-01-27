using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.9
    /// </summary>
    internal class ModuleReadOnly
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0025) { throw new ParseException("Failed to parse Id in ModuleReadOnly."); }

            // Reserved 
            position += 4;
        }
    }
}
