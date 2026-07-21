using System;
using System.Reflection;
Assembly asm = typeof(Avalonia.Input.Platform.IClipboard).Assembly;
foreach (Type t in asm.GetTypes())
{
    string n = t.Name;
    if (n.Contains("Clipboard") || n.Contains("DataTransfer") || n.Contains("DataFormat") || n.Contains("DataObject"))
    {
        Console.WriteLine("== " + t.FullName + (t.IsAbstract && t.IsSealed ? " [static]" : ""));
        foreach (MethodInfo m in t.GetMethods(BindingFlags.Public|BindingFlags.Static|BindingFlags.Instance|BindingFlags.DeclaredOnly))
            Console.WriteLine($"   {m.ReturnType.Name} {m.Name}({string.Join(", ", Array.ConvertAll(m.GetParameters(), p => p.ParameterType.Name))})");
    }
}
