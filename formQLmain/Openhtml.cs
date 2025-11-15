using System.Diagnostics;

namespace formQLmain
{
    internal class OpenHTML
    {
        private static readonly string defaultUrl = "https://trinhyenli.github.io/hdsd/";

        public static void OpenDefault()
        {
            Process.Start(new ProcessStartInfo(defaultUrl)
            {
                UseShellExecute = true
            });
        }
    }
}