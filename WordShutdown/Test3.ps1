# Saves and closes Word documents.
# Windows 11 64-bit. PowerShell 5.1

function wait
{
	param([int]$stop = 1)
	Start-Sleep -seconds $stop
}

add-type -AssemblyName microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")

$sig = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
Add-Type -MemberDefinition $sig -name NativeMethods -namespace Win32

for($i = 0; $i -lt 10; $i++)
{
	$test = Get-Process winword*;
	$last = tasklist /v /fi "WINDOWTITLE eq WORD" /fo list;

	$hwnd = @(Get-Process | Where-Object {$_.Name -eq "winword"})[0].MainWindowHandle
	[Win32.NativeMethods]::ShowWindowAsync($hwnd, 4)

	$stringy = 'INFORMACIÓN: no hay tareas ejecutándose que coincidan con los criterios especificados.'
	
	if ($test)
	{
		if ($last -like "*no hay tareas*")
		{
			$x = Get-Process winword
			foreach ($b in $x)
			{
				[Microsoft.VisualBasic.Interaction]::AppActivate($b.ID) *>$null
				[System.Windows.Forms.SendKeys]::SendWait("%{a}{d}")
				Start-Sleep -seconds 5
				[System.Windows.Forms.SendKeys]::SendWait("%{a}{e}")
			}
		}
		else
		{
			$x = Get-Process winword
			foreach ($b in $x)
			{
				taskkill /f /pid $b.ID
			}
		}
	}
}

exit
