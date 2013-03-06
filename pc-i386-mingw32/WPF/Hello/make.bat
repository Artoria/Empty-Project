@echo off
setlocal
set L=%windir%\Microsoft.NET\Framework\v4.0.30319\
set WPF=%L%WPF\
set REFWPF\=/r:%WPF%
set REFFX\=/r:%L%
set TARGET=/out:hello.exe 
set FILE=hello.cs
set REFERENCE=  %REFWPF\%PresentationFramework.dll  ^
  %REFWPF\%Windowsbase.dll                          ^
  %REFWPF\%Presentationcore.dll                     ^
  %REFWPF\%WindowsFormsIntegration.dll              ^
  %REFFX\%System.Xaml.dll             
set FLAGS=/nologo
csc %FLAGS% %TARGET% %REFERENCE% %FILE%
endlocal