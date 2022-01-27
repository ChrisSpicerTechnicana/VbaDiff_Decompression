using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects;
using VbaDiff.Decompression.Exceptions;
using VbaDiff.Model.References;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.2.1
    /// </summary>
    static internal class ReferenceReader
    {
        static internal RawReference ParseStream(byte[] stream, ref int position)
        {
            RawReferenceName referenceName;
            if (PeekForReferenceName(stream, position) == true)
            {
                referenceName = ReferenceNameReader.ParseStream(stream, ref position);
            }
            
            // Peek at this, don't read it
            uint referenceRecordType = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            
            ReferenceControl referenceControl;
            RawReference reference = null;
            switch (referenceRecordType)
            { 
                case 0x002F:
                    referenceControl = new ReferenceControl();
                    referenceControl.ParseStream(stream, ref position);
                    break;
                case 0x0033:
                    referenceControl = new ReferenceControl();
                    referenceControl.ParseStream(stream, ref position);
                    break;
                case 0x000D:
                    ReferenceRegistered referenceRegistered = new ReferenceRegistered();
                    referenceRegistered.ParseStream(stream, ref position);
                    break;
                case 0x000E:
                    reference = new RawReference();
                    reference.ReferenceProject = ReferenceProjectReader.ParseStream(stream, ref position);
                    break;                
                default:
                    throw new ParseException("Unknown referenceRecordType encountered!");                    
            }


            return reference;
        }

        static private bool PeekForReferenceName(byte[] stream, int position)
        {
            uint id = PeekNextUInt16(stream, position);

            return (id == 0x0016);
        }        

        static private uint PeekNextUInt16(byte[] stream, int position)
        {
            uint result = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);

            return result;
        }
    }
}
