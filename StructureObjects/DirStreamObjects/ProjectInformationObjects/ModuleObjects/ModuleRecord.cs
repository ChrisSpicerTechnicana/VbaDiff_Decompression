using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2
    /// </summary>
    class ModuleRecord
    {
        #region Fields
        internal ModuleName moduleName = new ModuleName();
        internal ModuleStreamName moduleStreamName = new ModuleStreamName();
        internal ModuleOffset moduleOffset;
        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream, ref int position)
        {
            moduleName.ParseStream(stream, ref position);
            ModuleNameUnicode moduleNameUnicode = new ModuleNameUnicode();
            moduleNameUnicode.ParseStream(stream, ref position);
            moduleStreamName.ParseStream(stream, ref position);
            ModuleDocString moduleDocString = new ModuleDocString();
            moduleDocString.ParseStream(stream, ref position);

            moduleOffset = new ModuleOffset();
            moduleOffset.ParseStream(stream, ref position);


            // Module Help Context, Module Cookie, Type, ReadOnly and Private add up to 40
            ModuleHelpContext moduleHelpContext = new ModuleHelpContext();
            moduleHelpContext.ParseStream(stream, ref position);
            ModuleCookie moduleCookie = new ModuleCookie();
            moduleCookie.ParseStream(stream, ref position);
            ModuleType moduleType = new ModuleType();
            moduleType.ParseStream(stream, ref position);

            if (PeekForModuleReadOnly(stream, position))
            {
                ModuleReadOnly moduleReadOnly = new ModuleReadOnly();
                moduleReadOnly.ParseStream(stream, ref position);
            }

            if(PeekForModulePrivate(stream, position))
            {
                position += 6;
            }

            // Terminator
            uint terminator = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (terminator != 0x002B) { throw new ParseException("Failed to parse terminator in ModuleRecord."); }

            // Reserved
            position += 4;
        }
        #endregion

        #region Private Methods
        private bool PeekForModuleReadOnly(byte[] stream, int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            return (id == 0x0025);
        }

        private bool PeekForModulePrivate(byte[] stream, int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);

            return (id == 0x0028);
        }
        #endregion

    }
}
