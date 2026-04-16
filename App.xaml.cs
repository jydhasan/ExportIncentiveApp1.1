using System;
using System.IO;
using System.Reflection;
using System.Runtime.Loader;

namespace ExportIncentiveApp
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(System.Windows.StartupEventArgs e)
        {
            // Fix: DinkToPdf needs its native library loaded explicitly
            // when running as a single-file EXE
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
            base.OnStartup(e);
        }

        private static Assembly? OnAssemblyResolve(object? sender, ResolveEventArgs args)
        {
            // Get the directory where the EXE is located
            string exeDir = AppContext.BaseDirectory;
            string assemblyName = new AssemblyName(args.Name).Name ?? "";
            string dllPath = Path.Combine(exeDir, assemblyName + ".dll");

            if (File.Exists(dllPath))
                return Assembly.LoadFrom(dllPath);

            return null;
        }
    }
}
