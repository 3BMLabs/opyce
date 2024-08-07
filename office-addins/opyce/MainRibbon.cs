using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace opyce
{

    public partial class MainRibbon
    {
        public static string tempPath = Path.GetTempPath();
        public static string opyceTempFolder = Path.Combine(tempPath, "opyce");
        public static string opyceIniFile = Path.Combine(tempPath, "opyce.ini");
        static FileSystemWatcher watcher = new FileSystemWatcher();
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public delegate string replaceFunction(string original);
        public static void SetPlaceHolders(string placeHolders)
        {
            //we cannot place it in the opyce directory, because we don't know if it's created yet
            string opyceIniPath = Path.Combine(Path.GetTempPath(), "opyce.ini");
            File.WriteAllText(opyceIniPath, placeHolders);
        }

        public void serialize<T>()

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
            if (Directory.Exists(opyceTempFolder))
            {
                try
                {
                    //delete previously created directory
                    Directory.Delete(opyceTempFolder, true);

                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.Message);
                    shouldCopy = false;
                }
            }
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
