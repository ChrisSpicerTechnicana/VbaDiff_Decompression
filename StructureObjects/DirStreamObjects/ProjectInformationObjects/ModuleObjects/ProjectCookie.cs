using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.1
    /// Project Cookies are always ignored, but they can form a useful check.
    /// </summary>
    class ProjectCookie
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0013) { throw new ParseException("Failed to parse ID in ProjectCookie."); }

            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000002) { throw new ParseException("Failed to parse size in ProjectCookie."); }

            // Cookie doesn't seem to be what it says it is in the book, so ignore it.            
            position += 2;            
        }
    }
}
