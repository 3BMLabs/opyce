using System;
using System.Diagnostics;

namespace opyce
{
    public class VSCodeRunner
    {
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
