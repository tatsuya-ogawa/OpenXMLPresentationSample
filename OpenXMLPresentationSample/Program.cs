// See https://aka.ms/new-console-template for more information

using System.Reflection;
using OpenXMLPresentationSample;

class Program
{
    public static void Main(string[] args)
    {
        var presentation = new PresentationEditor();
        var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var path = $"{assemblyPath}/Resources/Presentation.pptx";
        var newPath = $"{assemblyPath}/Resources/NewPresentation.pptx";
        var image = File.Open($"{assemblyPath}/Resources/earth.jpg",FileMode.Open);
        presentation.DoEdit(path,newPath,image);
    }
}