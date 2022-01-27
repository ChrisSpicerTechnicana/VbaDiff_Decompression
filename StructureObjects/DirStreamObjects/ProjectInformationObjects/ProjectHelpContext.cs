using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.8
    /// </summary>
    class ProjectHelpContext
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0007) { throw new ParseException("Failed to pass the Id for ProjectHelpContext."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000004) { throw new ParseException("Failed to parse the size for ProjectHelpContext."); }

            // Help Context... not interested in this.            
            position += 4;
        }
    }
}
