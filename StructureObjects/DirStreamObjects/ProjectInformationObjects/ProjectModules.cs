using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;
using VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.3
    /// </summary>
    class ProjectModules
    {
        #region Fields
        private List<ModuleRecord> moduleList = new List<ModuleRecord>();
        #endregion

        #region Properties
        internal List<ModuleRecord> Modulelist
        {
            get
            {
                return this.moduleList;
            }
        }

        #endregion

        internal void ParseStream(byte[] stream, ref int position)
        { 
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x000F) { throw new ParseException("Failed to parse ID in ProjectModules."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000002) { throw new ParseException("Failed to parse size in ProjectModules."); }

            // Count
            uint modulesCount = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            // Parse the Project Cookie. It's not used for anything but is a useful check.
            ProjectCookie projectCookie = new ProjectCookie();
            projectCookie.ParseStream(stream, ref position);

            // Modules
            uint counter = 0;
            while (counter < modulesCount)
            {
                ModuleRecord moduleRecord = new ModuleRecord();
                moduleRecord.ParseStream(stream, ref position);
                moduleList.Add(moduleRecord);
                counter++;
            }
        }
    }
}
