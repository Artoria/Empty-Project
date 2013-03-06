namespace Hello
{
    using System;
    using System.Windows;
    using System.Windows.Markup;
    using System.ComponentModel;
    using System.IO;
    using System.Xml;
    public class Test
    {
       [STAThread]
       public static void Main(String[] args){
         Window w;
         using(var stream = new StreamReader("hello.xaml")){
            w = (Window)XamlReader.Load(XmlReader.Create(new StringReader(stream.ReadToEnd())));
            w.Show();
         }
         Application app = new Application();
         app.Run();
       }
    }
   
}