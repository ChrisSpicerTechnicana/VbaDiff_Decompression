using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// Reads a ReferenceProject, see 2.3.4.2.2.6
    /// </summary>
    internal static class ReferenceProjectReader
    {        
        internal static RawReferenceProject ParseStream(byte[] stream, ref int position)
        {
            RawReferenceProject referenceProject = new RawReferenceProject();

            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x000E) { throw new ParseException("Failed to parse ID in ReferenceProject."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // SizeOfLibidAbsolute
            uint sizeOfLibidAbsolute = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in the LibIdAbsolute just yet.
            position += (int)sizeOfLibidAbsolute;

            // Size of LibRelative
            uint sizeOfLibRelative = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in the LibidRelative just yet.
            position += (int)sizeOfLibRelative;

            // Major Version
            uint majorVersion = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);            
            position += 4;

            // Minor version
            uint minorVersion = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            // Now check the size
            if (size != (4 + sizeOfLibidAbsolute + 4 + sizeOfLibRelative + 4 + 2)) { throw new ParseException("Size did not tally up in referenceProject."); }            

            return null;
        }

    }
}
