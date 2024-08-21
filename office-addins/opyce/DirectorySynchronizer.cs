using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace opyce
{
    public class DirectorySynchronizer
    {
        public string sourceFolder = "";
        public string destFolder = "";
        public FileSystemWatcher watcher;


        private void OnFileUpdate(object sender, FileSystemEventArgs e)
        {
            //copy file
            System.Threading.Thread.Sleep(1000);
            //sync
            string destFile = Path.Combine(destFolder, e.Name);
            string destDirectory = Path.GetDirectoryName(destFile);
            Directory.CreateDirectory(destDirectory);
            if (Directory.Exists(e.FullPath))
            {
                if (!Directory.Exists(destFile))
                {
                    //e is a directory, just create a new directory 
                    Directory.CreateDirectory(destFile);
                }
            }
            else if(File.Exists(e.FullPath))
            {
                //e is a file
                //copy the file to the folder
                File.Copy(e.FullPath, destFile, true);
            }
        }

        private void OnFileRename(object sender, RenamedEventArgs e)
        {
            if (Directory.Exists(Path.Combine(destFolder, e.OldName))){
                Directory.Move(Path.Combine(destFolder, e.OldName), Path.Combine(destFolder, e.Name));
            }
            else if(File.Exists(Path.Combine(destFolder, e.OldName)))
            {
                File.Move(Path.Combine(destFolder, e.OldName), Path.Combine(destFolder, e.Name));
            }
        }

        private void OnFileDelete(object sender, FileSystemEventArgs e)
        {
            string destFile = Path.Combine(destFolder, e.Name);
            if (Directory.Exists(destFile))
            {
                Directory.Delete(destFile, true);
            }
            else if (File.Exists(destFile))
            {
                File.Delete(destFile);
            }
        }

        public DirectorySynchronizer(string sourceFolder, string destFolder)
        {
            this.sourceFolder = sourceFolder;
            this.destFolder = destFolder;
            watcher = new FileSystemWatcher(sourceFolder, "*")
            {
                EnableRaisingEvents = true,
                IncludeSubdirectories = true
            };
            watcher.Created += OnFileUpdate;
            watcher.Changed += OnFileUpdate;
            watcher.Deleted += OnFileDelete;
            watcher.Renamed += OnFileRename;
        }
    }
}
