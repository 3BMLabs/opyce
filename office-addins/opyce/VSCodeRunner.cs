using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace opyce
{
    public class VSCodeRunner
    {
        public static void InstallVSCode()
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
        public static string RunCommand(string arguments)
        {
            try
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c {arguments}",//.Replace("\"", "\\\"")}",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process())
                {
                    process.StartInfo = processStartInfo;
                    process.Start();

                    // Capture both the standard output and standard error
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (!string.IsNullOrEmpty(error))
                    {
                        return $"Error: {error}";
                    }
                    return output;
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
}
