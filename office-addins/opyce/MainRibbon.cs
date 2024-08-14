using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Xml.Serialization;
using Office = Microsoft.Office.Core;
namespace opyce
{

    public partial class MainRibbon
    {
        public static string tempPath = Path.GetTempPath();
        public static string opyceTempFolder = Path.Combine(tempPath, "opyce");
        public static string opyceIniFile = Path.Combine(tempPath, "opyce.ini");
        static FileSystemWatcher watcher = new FileSystemWatcher();
        dynamic document;
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        void deleteOpyceFolder()
        {
            if (Directory.Exists(opyceTempFolder))
            {
                while (true)
                {
                    Process[] processes = Process.GetProcessesByName("Code");
                    if (processes.Length == 0) break;
                    if (MessageBox.Show("please close VS Code. Do you have the opyce folder opened? WE WILL DELETE THE FOLDER WHEN YOU CLICK NO", "Opyce", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        break;
                }
                try
                {
                    //delete previously created directory
                    Directory.Delete(opyceTempFolder, true);

                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void MainRibbon_Close(object sender, EventArgs e)
        {
            //stop watching files
            watcher = new FileSystemWatcher();
            deleteOpyceFolder();
        }

        public delegate string replaceFunction(string original);
        public static void SetPlaceHolders(string placeHolders)
        {
            //we cannot place it in the opyce directory, because we don't know if it's created yet
            string opyceIniPath = Path.Combine(Path.GetTempPath(), "opyce.ini");
            File.WriteAllText(opyceIniPath, placeHolders);
        }

        public void Serialize(dynamic document, bool write)
        {
            this.document = document;
            bool tempFolderExists = Directory.Exists(opyceTempFolder);
            DocumentProperties props = document.CustomDocumentProperties;
            if (write)
            {
                if (tempFolderExists)
                {
                    //delete all previous xml parts
                    foreach (DocumentProperty prop in props)
                    {
                        if (prop.Name.StartsWith("opyce"))
                        {
                            prop.Delete();
                        }
                    }

                    //save opyce folder within document structure

                    //loop over folders
                    string[] paths = Directory.GetFiles(opyceTempFolder, "*", SearchOption.AllDirectories);
                    int fileNumber = 0;
                    foreach (string path in paths)
                    {
                        string relativePath = path.Substring(opyceTempFolder.Length + 1);
                        props.Add("opyce file " + (fileNumber++), false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, relativePath + "\n" + File.ReadAllText(path));
                    }
                }
            }
            //we can't just overwrite the temp folder, because we'd merge random projects
            else
            {
                if (tempFolderExists)
                {
                    if(MessageBox.Show("opyce exists already in " + opyceTempFolder + ". Can we delete it?", "Opyce", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    {
                        deleteOpyceFolder();
                    }
                    else return;
                }
                Directory.CreateDirectory(opyceTempFolder);
                watcher = new FileSystemWatcher(opyceTempFolder, "*")
                {
                    EnableRaisingEvents = true,
                    IncludeSubdirectories = true
                };
                watcher.Created += OnFileUpdate;
                watcher.Changed += OnFileUpdate;
                foreach(DocumentProperty prop in props)
                {
                    if (prop.Name.StartsWith("opyce"))
                    {
                        int separatorIndex = prop.Value.IndexOf('\n');
                        
                        string value = prop.Value as string;

                        string absolutePath = Path.Combine(opyceTempFolder, value.Substring(0, separatorIndex));
                        Directory.CreateDirectory(Path.GetDirectoryName(absolutePath));
                        File.WriteAllText(absolutePath, value.Substring(separatorIndex + 1));
                    }
                }
                CustomXML[] customData = Serializer.GetCustomXmlParts<CustomXML>(document, Serializer.OpyceNameSpace);
                foreach (CustomXML data in customData)
                {
                    string absolutePath = Path.Combine(opyceTempFolder, data.Key);
                    Directory.CreateDirectory(Path.GetDirectoryName(absolutePath));
                    File.WriteAllText(absolutePath, data.Value);
                }
            }
        }

        private void OnFileUpdate(object sender, FileSystemEventArgs e)
        {
            document.Saved = false;
        }

        public void InstallVSCode()
        {
            using (var client = new WebClient())
            {
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    string vscodeInstallerUrl = "https://update.code.visualstudio.com/latest/win32-x64-user/stable";

                    string tempFilePath = Path.Combine(Path.GetTempPath(), "VSCodeSetup.exe");
                    if (!File.Exists(tempFilePath) ||
                        MessageBox.Show("The installer has been downloaded already to '" + tempFilePath + "'. do you want to redownload it?", "Opyce", MessageBoxButtons.OKCancel) == DialogResult.OK)
                        client.DownloadFile(vscodeInstallerUrl, tempFilePath);

                    if (MessageBox.Show("VS Code finished downloading and will be installed now. IMPORTANT: RESTART YOUR PC AFTER INSTALLATION!", "Opyce", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    // Run the installer.
                    {
                        Process installationProcess = Process.Start(tempFilePath);
                        void process_Exited(object sender, EventArgs e)
                        {
                            MessageBox.Show("Don't forget to restart your PC!", "Opyce", MessageBoxButtons.OK);

                            // do something when process terminates;
                        }
                        installationProcess.Exited += process_Exited;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while trying to install VS Code: " + ex.Message);
                }
            }
        }

        private void OpenInPythonButton_Click(object sender, RibbonControlEventArgs e)
        {
            bool shouldCopy = true;
            //create a folder, containing the code files
            //if (Directory.Exists(opyceTempFolder))
            //{
            //    //check if the ini file contains
            //    MessageBox.Show("a folder already exists in " + opyceTempFolder);
            //    return;
            //}
            if (shouldCopy)
            {
                FileUtils.CopyDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"clone-folder"), opyceTempFolder, true, false);
            }

            string backendFolderPath = Path.Combine(opyceTempFolder, "backend");
            string backendFilePath = Path.Combine(backendFolderPath, "main.py");

            if (File.Exists(backendFilePath))
            {
                string backendFileText = File.ReadAllText(backendFilePath);

                foreach (string line in File.ReadLines(opyceIniFile))
                {
                    int splitIndex = line.IndexOf('=');
                    string tag = line.Substring(0, splitIndex);
                    string placeHolder = line.Substring(splitIndex + 1);
                    backendFileText = backendFileText.Replace("$" + tag + "$", placeHolder);
                }

                File.WriteAllText(backendFilePath, backendFileText);
            }
            //string vsCodePath = Path.Combine(opyceTempFolder, ".vscode");
            //Directory.CreateDirectory(vsCodePath);
            //
            ////add extension recommendations
            //string recommendationFile = Path.Combine(vsCodePath, "extensions.json");
            //string[] recommendations = { "ms-python.python" };
            //const string StartOfLine = "\n        \"";
            //const string EndOfLine = "\",";
            //string recommendationString = "{\n    \"recommendations\": [" + StartOfLine + String.Join(StartOfLine + EndOfLine, recommendations) + EndOfLine + "\n    ]\n}";
            //File.WriteAllText(recommendationFile, recommendationString);
            ////make vscode open module1
            //
            //var SubFolderPath = Path.Combine(opyceTempFolder, "module1");
            //Directory.CreateDirectory(SubFolderPath);
            //File.WriteAllText(Path.Combine(SubFolderPath, "main.py"), "#if you have not installed python:\n#\tgo to the run section (click on the triangle bug icon on the left)\n#\tclick 'Run and Debug'\n#\t follow instructions.\n");

            //arguments:
            //. -> open current folder
            //module1/main.py -> open desired file in folder
            string vsoutput = VSCodeRunner.RunCommand($"cd \"{opyceTempFolder}\" && code . module1/main.py");


            if (vsoutput.StartsWith("Error"))
            {
                if (MessageBox.Show($"We encountered an error trying to launch vscode:\n{vsoutput}\nIt appears that Visual Studio Code (the python editor) has not been installed.\nDo you want to install it?", "Opyce", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    InstallVSCode();
                }
            }

            //using (var client = new HttpClient())
            //{
            //    using (var s = client.GetStreamAsync("https://via.placeholder.com/150"))
            //    {
            //        using (var fs = new FileStream("localfile.jpg", FileMode.OpenOrCreate))
            //        {
            //            s.Result.CopyTo(fs);
            //        }
            //    }
            //}
        }

    }
}
