using Grasshopper.Kernel;
using System;
using System.IO;
using System.Reflection;

namespace MyGrasshopperAssembly1
{
    public class ReflectionFrom
    {
        public void InvokeMethod(IGH_DataAccess DA, string assembleName, string FullClassName, string methodName)
        {
            string mainAssemblyPath = Path.Combine(HardCodedData.Bim_directoryPath, assembleName);
            byte[] dllBytes = null;

            try
            {
                dllBytes = File.ReadAllBytes(mainAssemblyPath);
            }
            catch (Exception ex)
            {
                //SendEmailNotification("DLL File Reading Error", $"Error reading DLL file: {ex.Message}");
                return;
            }
            Assembly mainAssembly = Assembly.Load(dllBytes);

            // Set up the assembly resolver to handle dependent assemblies
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;


            string className = FullClassName;
            Type type = mainAssembly.GetType(className);
            object instance = Activator.CreateInstance(type);

            // Assuming Execute method has parameters like (object arg1, ref string arg2, List<object> arg3)
            object[] methodParameters = new object[] { DA };
            // Getting MethodInfo for Execute method
            MethodInfo executeMethod = type.GetMethod(methodName);
            // Invoking the Execute method on the instance
            executeMethod.Invoke(instance, methodParameters);
        }
        public Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var name = new AssemblyName(args.Name).Name;

            // Construct the full path to the missing assembly in the main assembly's directory
            string missingAssemblyPath = HardCodedData.Bim_directoryPath + "\\" + name + ".dll";

            // Check if the dependent assembly exists in the main assembly's directory
            if (File.Exists(missingAssemblyPath))
            {
                // Load and return the dependent assembly
                byte[] dependentAssemblyBytes = File.ReadAllBytes(missingAssemblyPath);
                return Assembly.Load(dependentAssemblyBytes);
            }

            return null;
        }
    }
}
