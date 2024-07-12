using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaskLaunchTest
{
    [Guid("86C5173C-AD05-470D-AA08-59D4EF81FEDC"),ComVisible(true)]
    public class Class1:IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "TaskLaunchTest";
            poInfo.mbsCompany = "Company";
            poInfo.mbsDescription = "TaskLaunchTest";
            poInfo.mlAddInVersion = 1;
            poInfo.mlRequiredVersionMajor = 10;
            poInfo.mlRequiredVersionMinor = 0;

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskLaunch);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
        }
        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            switch (poCmd.meCmdType)
            {
                case EdmCmdType.EdmCmd_TaskLaunch:
                    OnTaskLaunch(ref poCmd, ref ppoData);
                    break;
                case EdmCmdType.EdmCmd_TaskSetup:
                    OnTaskSetup(ref poCmd, ref ppoData);
                    break;
            }
        }
        private void OnTaskSetup(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            IEdmTaskProperties taskProperties = (IEdmTaskProperties)poCmd.mpoExtra;
            taskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitExec;
        }
        private void OnTaskLaunch(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            IEdmTaskInstance taskInstance = (IEdmTaskInstance)poCmd.mpoExtra;
            TaskLaunchControl taskLaunchControl = new TaskLaunchControl();
            taskLaunchControl.CreateControl();
            poCmd.mbsComment = "TaskLaunchTest";
            poCmd.mlParentWnd = taskLaunchControl.Handle.ToInt32();
            poCmd.mpoExtra = taskInstance;
        }
    }
}
