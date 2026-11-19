using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

internal static class Launcher
{
    [STAThread]
    private static int Main()
    {
        string scriptPath = Path.Combine(Path.GetTempPath(), "WebsitesDoIt-LeadFinder-" + Guid.NewGuid().ToString("N") + ".ps1");
        StringBuilder diagnostics = new StringBuilder();
        try
        {
            using (Stream script = Assembly.GetExecutingAssembly().GetManifestResourceStream("LeadFinder.ps1"))
            {
                if (script == null) throw new InvalidOperationException("The application script was not bundled.");
                using (FileStream output = File.Create(scriptPath)) script.CopyTo(output);
            }

            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = powershell;
            start.Arguments = "-NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File \"" + scriptPath + "\"";
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.WindowStyle = ProcessWindowStyle.Hidden;
            start.RedirectStandardError = true;
            using (Process process = Process.Start(start))
            {
                process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e)
                {
                    if (e.Data != null) diagnostics.AppendLine(e.Data);
                };
                process.BeginErrorReadLine();
                process.WaitForExit();
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    string detail = diagnostics.ToString().Trim();
                    if (String.IsNullOrEmpty(detail)) detail = "The app could not start. Try rebuilding or contact support.";
                    MessageBox.Show(detail, "WebsitesDoIt Lead Finder could not start", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return process.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "WebsitesDoIt Lead Finder could not start", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
        finally
        {
            try { if (File.Exists(scriptPath)) File.Delete(scriptPath); } catch { }
        }
    }
}
