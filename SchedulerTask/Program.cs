using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.Linq;


namespace SchedulerTask
{
    class Program: SchedulerTaskBus
    {
        static void Main(string[] args)
        {
            //CreateOnceRunTask("Test","hello12","nihao","gg",DateTime.Now,"");
            DeleteSubFolders("\\Test\\qwe");
            string str = "\\Test";
            List<TaskFolder> jj = GetAllSubFolder(str).ToList();
            List<Task> hh = GetAllTask().ToList();
            List<Task> gg =GetAllTask(str);
            Task task=GetTask("hello");
            bool b = IsExist("Test","hello");
            Console.ReadKey();
        }
    }
}
