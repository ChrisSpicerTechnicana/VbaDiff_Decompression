using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.2
    /// </summary>
    internal class ModuleNameUnicode
    {
        internal void ParseStream(byte[] stream, ref int position)
        { 
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0047) { throw new ParseException("Failed to Parse ID in ModuleNameUnicode."); }

            // SizeOfModuleNameUnicode
            uint sizeOfModuleNameUnicode = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if ((sizeOfModuleNameUnicode % 2) != 0) { throw new ParseException("Failed to parse sizein ModuleNameUnicode."); }

            // Not interested in ModuleNameUnicode for V1?
            position += (int) sizeOfModuleNameUnicode;
        }
    }
}
