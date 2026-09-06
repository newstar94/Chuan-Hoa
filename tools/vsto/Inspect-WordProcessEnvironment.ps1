param([Parameter(Mandatory=$true)][int]$WordProcessId)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WordEnvironmentProbe {
 [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
 [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
 [DllImport("kernel32.dll", SetLastError=true)] static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] data, int size, out IntPtr read);
 [DllImport("ntdll.dll")] static extern int NtQueryInformationProcess(IntPtr process, int info, IntPtr[] data, int size, out int returned);
 static byte[] Read(IntPtr handle, long address, int length) {
  var data=new byte[length]; IntPtr count;
  if(!ReadProcessMemory(handle,new IntPtr(address),data,length,out count) || count.ToInt64()!=length)
   throw new InvalidOperationException("Cannot read process metadata.");
  return data;
 }
 public static string ReadEnvironment(int pid) {
  if(IntPtr.Size!=8) throw new InvalidOperationException("Use 64-bit PowerShell.");
  var handle=OpenProcess(0x410,false,pid);
  if(handle==IntPtr.Zero) throw new InvalidOperationException("Cannot open process metadata.");
  try {
   var info=new IntPtr[6]; int returned;
   if(NtQueryInformationProcess(handle,0,info,48,out returned)!=0) throw new InvalidOperationException("Cannot query PEB.");
   long parameters=BitConverter.ToInt64(Read(handle,info[1].ToInt64()+0x20,8),0);
   long environment=BitConverter.ToInt64(Read(handle,parameters+0x80,8),0);
   var text=new StringBuilder();
   for(int offset=0;offset<131072;offset+=2) {
    char c=(char)BitConverter.ToUInt16(Read(handle,environment+offset,2),0);
    if(c==0 && text.Length>0 && text[text.Length-1]==0) break;
    text.Append(c);
   }
   return text.ToString();
  } finally {CloseHandle(handle);}
 }
}
'@
$allowed = '^(VSTO|COMPLUS|COR_|DOTNET_|__COMPAT_LAYER$|APPDATA$|LOCALAPPDATA$|USERPROFILE$|TEMP$|TMP$)'
[WordEnvironmentProbe]::ReadEnvironment($WordProcessId).Split([char]0) | ForEach-Object {
 $parts=$_.Split('=',2)
 if($parts.Length -eq 2 -and $parts[0] -match $allowed) {
  [pscustomobject]@{Name=$parts[0]; WordValue=$parts[1]; DiagnosticProcessValue=[Environment]::GetEnvironmentVariable($parts[0])}
 }
} | Format-List
