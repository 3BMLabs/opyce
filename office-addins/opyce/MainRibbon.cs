using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using Office = Microsoft.Office.Core;
namespace opyce
{

    public partial class MainRibbon
    {
        //variables initialized by both applications
        public static string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        public static string opyceAppDataFolder = Path.Combine(appDataPath, "opyce");
        public static string tempPath = Path.GetTempPath();
        public static string opyceTempFolder = Path.Combine(tempPath, "opyce");
        public static string opyceIniFile = Path.Combine(tempPath, "opyce.ini");

        //variables unknown by the mainribbon
        public static string appName = "";

        dynamic document;

        public static DirectorySynchronizer synchronizer = null;
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
            synchronizer = null;
            deleteOpyceFolder();
        }

        public delegate string replaceFunction(string original);
        public static void SetPlaceHolders(string appName, string initialization = "", string placeHolders = "")
        {
            MainRibbon.appName = appName;
            placeHolders = $"appname={appName}\ninitialization={initialization}\n" + placeHolders;
            //we cannot place it in the opyce directory, because we don't know if it's created yet
            string opyceIniPath = Path.Combine(Path.GetTempPath(), "opyce.ini");
            File.WriteAllText(opyceIniPath, placeHolders);
        }

        public void Serialize(bool write, dynamic document = null)
        {
            this.document = document;
            bool tempFolderExists = Directory.Exists(opyceTempFolder);
            string serializedFolderPath = Path.Combine(MainRibbon.opyceAppDataFolder, appName);
            if (write)
            {
                if (tempFolderExists)
                {
                    if (document != null)
                    {
                        //save in document

                        //delete all previous xml parts
                        foreach (DocumentProperty prop in document.CustomDocumentProperties)
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
                            document.CustomDocumentProperties.Add("opyce file " + (fileNumber++), false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, relativePath + "\n" + File.ReadAllText(path));
                        }
                    }
                    else
                    {
                        //copy to appdata folder
                        FileUtils.CopyDirectory(opyceTempFolder, serializedFolderPath, true, true);
                    }
                }
            }
            //we can't just overwrite the temp folder, because we'd merge random projects
            else
            {
                if (tempFolderExists)
                {
                    if (MessageBox.Show("opyce exists already in " + opyceTempFolder + ". Can we delete it?", "Opyce", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    {
                        deleteOpyceFolder();
                    }
                    else return;
                }

                if (document != null)
                {
                    Directory.CreateDirectory(opyceTempFolder);
                    //load from document
                    foreach (DocumentProperty prop in document.CustomDocumentProperties)
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
                else
                {
                    if (Directory.Exists(serializedFolderPath))
                    {
                        //copy from appdata folder
                        FileUtils.CopyDirectory(serializedFolderPath, opyceTempFolder, true, false);
                    }
                    else
                    {
                        Directory.CreateDirectory(opyceTempFolder);
                    }
                    synchronizer = new DirectorySynchronizer(opyceTempFolder, serializedFolderPath);
                    synchronizer.watcher.Changed += OnFileUpdate;
                    synchronizer.watcher.Deleted += OnFileUpdate;
                    synchronizer.watcher.Created += OnFileUpdate;
                    synchronizer.watcher.Renamed += OnFileUpdate;
                }

            }
        }

        string[][] readIniFile(string path)
        {
            string[] lines = File.ReadAllLines(path);
            string[][] pairs = new string[lines.Length][];
            int lineIndex = 0;
            foreach (string line in lines)
            {
                string trimmedLine = line.Split(';')[0];
                int splitIndex = line.IndexOf('=');
                if (splitIndex == -1) { continue; }
                string name = line.Substring(0, splitIndex);
                string value = line.Substring(splitIndex + 1);
                pairs[lineIndex++] = new string[2] { name, value };
            }
            return pairs.Where(pair => pair != null).ToArray();
        }

        private void OnFileUpdate(object sender, FileSystemEventArgs e)
        {

            if (document != null)
                document.Saved = false;
            //we support max. 9 files, since buttons can't be added dynamically.

            if (e.Name == "backend\\buttons.ini")
            {
                System.Threading.Thread.Sleep(1000);
                //read out files
                string buttonFile = Path.Combine(opyceTempFolder, "backend", "buttons.ini");

                string[][] pairs = File.Exists(buttonFile) ? readIniFile(buttonFile) : new string[0][];
                RibbonButton[] buttons = { pyfunc1, pyfunc2, pyfunc3, pyfunc4, pyfunc5, pyfunc6, pyfunc7, pyfunc8, pyfunc9 };
                int buttonIndex = 0;
                foreach (string[] pair in pairs)
                {
                    buttons[buttonIndex].Label = pair[0];
                    buttons[buttonIndex].Visible = true;
                    buttons[buttonIndex].Tag = pair[1];
                    buttonIndex++;
                }
                for (; buttonIndex < buttons.Length; buttonIndex++)
                {
                    buttons[buttonIndex].Visible = false;
                }
            }
        }

        private void executePythonFunction(object sender, RibbonControlEventArgs e)
        {
            RibbonButton button = (RibbonButton)sender;
            string command = $"cd \"{opyceTempFolder}\" && set \"PYTHONPATH={opyceTempFolder}\" && python {button.Tag}";
            Task.Run(() =>
            {
                string vsoutput = VSCodeRunner.RunCommand(command);
            });

        }

        private void OpenInPythonButton_Click(object sender, RibbonControlEventArgs e)
        {
            openInPython.Label = "blah";
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
                    VSCodeRunner.InstallVSCode();
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
