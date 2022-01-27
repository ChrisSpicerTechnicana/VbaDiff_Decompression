using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects;
using VbaDiff.Decompression.Exceptions;
using System.Diagnostics;
using VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects;
using VbaDiff;
using Technicana.Infrastructure.Utilities.Strings;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects
{

    /// <summary>
    /// 2.3.4.2.1
    /// </summary>
    class ProjectInformation
    {
        #region Fields
        private ProjectSysKind sysKindRecord = new ProjectSysKind();
        private ProjectLcid lcidRecord = new ProjectLcid();
        private ProjectLcidInvoke lcidInvokeRecord = new ProjectLcidInvoke();
        private ProjectCodePage projectCodePage = new ProjectCodePage();
        private ProjectNameRecord projectNameRecord = new ProjectNameRecord();
        private ProjectDocString projectDocString = new ProjectDocString();
        private ProjectHelpFilePath projectHelpFilePath = new ProjectHelpFilePath();
        private ProjectLibFlagsRecord projectLibFlagsRecord = new ProjectLibFlagsRecord();
        private ProjectVersion projectVersion = new ProjectVersion();
        private ProjectConstants projectConstants = new ProjectConstants();
        private ProjectHelpContext projectHelpContext = new ProjectHelpContext();
        private ProjectReferences projectReferences = new ProjectReferences();
        private ProjectModules projectModules = new ProjectModules();

        // Result Variables
        private UInt16 codePage;

        private Dictionary<string, int> moduleOffsets = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Properties
        internal UInt16 CodePage
        {
            get
            {
                return this.codePage;   
            }
        }

        internal Dictionary<string, int> ModuleOffsets
        {
            get
            {
                return this.moduleOffsets;
            }
        }

        internal ProjectNameRecord ProjectNameRecord
        {
            get
            {
                return this.projectNameRecord;
            }
        }

        internal ProjectModules ProjectModules
        {
            get
            {
                return this.projectModules;
            }
        }

        /// <summary>
        /// Returns ProjectConstants Object, representing the Project Constants set on this project.
        /// </summary>
        internal ProjectConstants ProjectConstants
        {
            get
            {
                return this.projectConstants;
            }
        }

        internal ProjectReferences ProjectReferences
        {
            get
            {
                return this.projectReferences;
            }
        }
        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream, ref int position)
        {
            try
            {
                sysKindRecord.ParseStream(stream, ref position);
                lcidRecord.ParseStream(stream, ref position);
                lcidInvokeRecord.ParseStream(stream, ref position);
                
                // This parsing is particularly important as it gives you the code page with which to decode project strings.
                projectCodePage.ParseStream(stream, ref position);
                this.codePage = projectCodePage.CodePage;

                projectNameRecord.ParseStream(stream, ref position);

                //// We should have enough information now to get the Project name.
                //this.projectName = DecodeString(projectNameRecord.ProjectName);

                projectDocString.ParseStream(stream, ref position);
                projectHelpFilePath.ParseStream(stream, ref position);
                projectHelpContext.ParseStream(stream, ref position);
                projectLibFlagsRecord.ParseStream(stream, ref position);
                projectVersion.ParseStream(stream, ref position);
                projectConstants.ParseStream(stream, ref position);

                // Now parse the Project References. You basically keep parsing references until you reach a Project Module.
                projectReferences.ParseStream(stream, ref position);

                // Now parse the Project Modules Record.
                projectModules.ParseStream(stream, ref position);

                foreach (ModuleRecord moduleRecord in projectModules.Modulelist)
                {
                    string moduleName = StringUtilities.DecodeString(moduleRecord.moduleName.ModuleNameData, this.codePage);                    
                    moduleOffsets.Add(moduleName, (int) moduleRecord.moduleOffset.TextOffset);
                }
            }
            catch (ParseException e)
            {
                throw new ParseException("Could not properly parse the ProjectInformation Record. Check inner exception for why.", e);                 
            }
        }
        #endregion


    }
}
