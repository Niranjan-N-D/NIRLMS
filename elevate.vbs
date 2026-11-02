Set UAC = CreateObject("Shell.Application")
UAC.ShellExecute "cmd.exe", "/c """ & "C:\Projects\library-main (1)\library-main1\apply_update.cmd" & """", "", "runas", 1
