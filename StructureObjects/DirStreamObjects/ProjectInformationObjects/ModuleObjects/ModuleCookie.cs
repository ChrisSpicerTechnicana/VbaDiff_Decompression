using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.7
    /// </summary>
    internal class ModuleCookie
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x002c) { throw new ParseException("Failed to parse ID in ModuleCookie."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000002) { throw new ParseException("Failed to parse Size in ModuleCookie."); }
            
            // Ignore Cookie
            position += 2;
        }
    }
}
