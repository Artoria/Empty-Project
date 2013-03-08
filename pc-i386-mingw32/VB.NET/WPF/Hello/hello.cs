Imports System
Imports System.Windows
Imports System.Windows.Markup
Imports System.ComponentModel
Imports System.IO
Imports System.Xml
Namespace Hello
    Public class Test
       Public Shared Sub Main(args as String())
         Dim w as Window, stream as StreamReader = new StreamReader("hello.xaml")
         Dim app as Application = new Application
         w = XamlReader.Load(XmlReader.Create(new StringReader(stream.ReadToEnd())))
         w.Show()
         stream.Close()
         app.Run()
       End Sub
    End Class
End Namespace