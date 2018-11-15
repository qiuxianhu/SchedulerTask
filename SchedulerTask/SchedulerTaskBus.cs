using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SchedulerTask
{
    public class SchedulerTaskBus
    {
        /// <summary>
        /// 创建一个以指定日期时间为触发条件的计划任务
        /// </summary>
        /// <param name="subFolderName">计划任务所在的文件夹</param>
        /// <param name="taskName">计划任务的名称</param>
        /// <param name="taskDescription">计划任务描述</param>
        /// <param name="path">应用程序路径</param>
        /// <param name="startBoundary">指定日期时间</param>
        /// <param name="arguments">应用程序参数</param>       
        /// 
        ///using语句，定义一个范围，在范围结束时处理对象。 场景：
        ///当在某个代码段中使用了类的实例，而希望无论因为什么原因，只要离开了这个代码段就自动调用这个类实例的Dispose。
        ///要达到这样的目的，用try...catch来捕捉异常也是可以的，但用using也很方便。 
        public static void CreateOnceRunTask(string subFolderName, string taskName, string taskDescription, string path, DateTime startBoundary, string arguments)
        {
            //以指定日期時間为触发时间初始化触发器
            using (TimeTrigger trigger = new TimeTrigger(startBoundary))
            {
                ///初始化执行任务的应用程序
                using (ExecAction action = new ExecAction(path, arguments, null))
                {
                    CreateTask(subFolderName, taskName, taskDescription, trigger, action);
                }
            }
        }

        /// <summary>
        /// 创建一个计划任务
        /// </summary>
        /// <param name="taskName">计划任务名称</param>
        /// <param name="taskDescription">计划任务描述</param>
        /// <param name="trigger">触发条件</param>
        /// <param name="action">执行任务</param>
        /// <param name="subFolderName">计划任务所在的文件夹</param>
        public static void CreateTask(string subFolderName, string taskName, string taskDescription, Trigger trigger, Microsoft.Win32.TaskScheduler.Action action)
        {
            using (TaskService ts = new TaskService())
            {
                using (TaskDefinition td = ts.NewTask())
                {
                    td.RegistrationInfo.Description = taskDescription;
                    td.RegistrationInfo.Author = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                    #region LogonType元素和UserId元素被一起使用，以定义需要运行这些任务的用户。
                    //http://stackoverflow.com/questions/8170236/how-to-set-run-only-if-logged-in-and-run-as-with-taskscheduler-in-c
                    //td.Principal.UserId = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    //td.Principal.LogonType = TaskLogonType.InteractiveToken;
                    #endregion
                    td.Principal.RunLevel = TaskRunLevel.Highest;
                    td.Triggers.Add(trigger);
                    td.Actions.Add(action);
                    TaskFolder folder = ts.RootFolder;
                    if (!string.IsNullOrWhiteSpace(subFolderName))
                    {
                        if (!folder.SubFolders.Exists(subFolderName))
                        {
                            folder = folder.CreateFolder(subFolderName);
                        }
                        else
                        {
                            folder = folder.SubFolders[subFolderName];
                        }
                    }
                    folder.RegisterTaskDefinition(taskName, td);
                    folder.Dispose();
                    folder = null;
                }
            }
        }


        /// <summary>
        /// 删除某一个子文件夹
        /// </summary>
        /// <param name="folderName">绝对路径</param>
        public static void DeleteSubFolders(string folderName = @"\Miscrosoft")
        {
            // Get the service on the local machine 
            using (TaskService taskService = new TaskService())
            {
                using (TaskFolder folder = taskService.GetFolder(folderName))
                {
                    if (folder.Tasks.Count > 0)
                    {
                        throw new InvalidOperationException("TASKS_EXIST");
                    }
                    if (folder.SubFolders.Count > 0)
                    {
                        throw new InvalidOperationException("SUB_FOLDERS_EXIST");
                    }
                    using (TaskFolder folderParent = taskService.GetFolder(folder.Path.Replace(@"\" + folder.Name, "")))
                    {
                        folderParent.DeleteFolder(folder.Name);
                    }
                }
            }
        }

        /// <summary>
        /// 获取某一个文件夹下面的所有子文件夹
        /// </summary>
        /// <param name="folderName">传null表示根文件夹</param>
        /// <returns></returns>
        public static IEnumerable<TaskFolder> GetAllSubFolder(string folderName = @"\Miscrosoft")
        {
            using (TaskService taskService = new TaskService())
            {
                TaskFolder folder = taskService.RootFolder;
                if (!string.IsNullOrWhiteSpace(folderName))
                {
                    folder = taskService.GetFolder(folderName);
                }
                foreach (TaskFolder subFolder in folder.SubFolders)
                {
                    yield return subFolder;
                }
            }
        }

        /// <summary>
        /// 删除某一个文件夹下面的某一个任务
        /// </summary>
        /// <param name="taskName">要删除的任务名称</param>
        /// <param name="folderName">传null表示根文件夹</param>
        public static void DeleteTask(string folderName, string taskName)
        {
            using (TaskService taskService = new TaskService())
            {
                TaskFolder folder = taskService.RootFolder;
                // 判断文件夹和任务是否存在
                if (folder.SubFolders.Exists(folderName) && folder.SubFolders[folderName].Tasks.Exists(taskName))
                {
                    folder.SubFolders[folderName].DeleteTask(taskName);
                }
                folder.Dispose();
                folder = null;
            }
        }

        /// <summary>
        /// 获取所有任务
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<Task> GetAllTask()
        {
            using (TaskService taskService = new TaskService())
            {
                return taskService.AllTasks;
            }
        }

        /// <summary>
        /// 获取某一个文件夹下面的所有任务
        /// </summary>
        /// <param name="folderName">传null表示根文件夹</param>
        /// <returns></returns>
        public static List<Task> GetAllTask(string folderName = @"\Miscrosoft")
        {
            using (TaskService taskService = new TaskService())
            {
                TaskFolder folder = taskService.RootFolder;
                if (!string.IsNullOrWhiteSpace(folderName) && folder.SubFolders.Exists(folderName))
                {
                    folder = taskService.GetFolder(folderName);
                    return folder.AllTasks.ToList();
                }
                return new List<Task>(0);
            }
        }

        /// <summary>
        /// 查找符合条件的所有Task
        /// </summary>
        /// <param name="regex"></param>
        /// <param name="searchAllFolders"></param>
        /// <returns></returns>
        public static IEnumerable<Task> GetAllTask(Regex regex, bool searchAllFolders = true)
        {
            using (TaskService taskService = new TaskService())
            {
                return taskService.FindAllTasks(regex, searchAllFolders);
            }
        }

        /// <summary>
        /// 查找符合条件的Task
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="searchAllFolders"></param>
        /// <returns></returns>
        public static Task GetTask(string taskName, bool searchAllFolders = true)
        {
            using (TaskService taskService = new TaskService())
            {
                return taskService.FindTask(taskName, searchAllFolders);
            }
        }

        public static bool IsExist(string folderName, string taskName)
        {
            using (TaskService taskService = new TaskService())
            {
                TaskFolder folder = taskService.RootFolder;
                // 判断文件夹和任务是否存在
                return folder.SubFolders.Exists(folderName) && folder.SubFolders[folderName].Tasks.Exists(taskName);
            }
        }
    }
}
