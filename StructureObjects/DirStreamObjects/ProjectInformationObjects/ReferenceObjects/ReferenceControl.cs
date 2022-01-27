using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// 2.3.4.2.2.3
    /// </summary>
    class ReferenceControl
    {
        internal void ParseStream(byte[] stream, ref int position)
        {
            // The ReferenceOrginal record should exist, but may not. 
            if (PeekForOriginalRecord(stream, position))
            {
                ReferenceOriginal referenceOriginal = new ReferenceOriginal();
                referenceOriginal.ParseStream(stream, ref position);
            }

            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x002F) { throw new ParseException("Failed to parse ID from ReferenceControl."); }

            // SizeTwiddled
            uint sizeTwiddled = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            uint sizeOfLibidTwiddled = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in LibidTwiddled
            position += (int)sizeOfLibidTwiddled;

            // Check Reserved 1 & 2
            uint reserved1 = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (reserved1 != 0x00000000) { throw new ParseException("Failed to parse Reserved1 from ReferenceControl."); }

            uint reserved2 = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved2 != 0x0000) { throw new ParseException("Failed to parse Reserved2 from ReferenceControl."); }

            // Check for NameRecordExtended. This field is optional.
            if (PeekForReferenceNameRecord(stream, position))
            {
                ReferenceNameReader.ParseStream(stream, ref position);
            }

            // Reserved 3
            uint reserved3 = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved3 != 0x0030) { throw new ParseException("Failed to parse Reserved3 from ReferenceControl."); }

            // SizeExtended. Not so interested in this.
            position += 4;

            // SizeOfLibidExtended
            uint sizeOfLibidExtended = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            // Not interested in LibIdExtended
            position += (int)sizeOfLibidExtended;

            // Reserved 4
            uint reserved4 = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (reserved4 != 0x00000000) { throw new ParseException("Failed to parse Reserved4 from ReferenceControl."); }

            // Reserved 5
            uint reserved5 = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            // OriginalTypeLib
            position += 16;

            // Cookie
            position += 4;           
        }

        private bool PeekForOriginalRecord(byte[] stream, int position)
        { 
            uint nextBytes = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);

            return (nextBytes == 0x0033);
        }

        private bool PeekForReferenceNameRecord(byte[] stream, int position)
        {
            uint nextBytes = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);

            return (nextBytes == 0x0016);
        }
    }
}
