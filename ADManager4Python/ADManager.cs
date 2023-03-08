using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace ADM4P
{
    public class ADManager4Python : AppDomainManager
    {
        #region Activation context

        const int constWorkingDir = 0;
        const int constBaseDir = 1;
        const int constAppName = 2;
        const int constBinPath = 3;
        const int constCfgFile = 4;
        const int constTgtFramework = 5;
        private static string[] s_ActivationContext;

        #endregion

        #region Entry assembly

        private static Assembly s_AsmMscorlib;
        public override Assembly EntryAssembly
        {
            get
            {
                return s_AsmMscorlib;
            }
        }

        #endregion

        public override void InitializeNewDomain(AppDomainSetup appDomainInfo)
        {
            if (AppDomain.CurrentDomain.IsDefaultAppDomain())
            {
                // read activation context from environmet variable
                string activationContext = System.Environment.GetEnvironmentVariable("ADM_4_PYTHON_ACTIVATION_CONTEXT");
                if (!string.IsNullOrEmpty(activationContext))
                {
                    s_AsmMscorlib = typeof(Object).Assembly;
                    s_ActivationContext = activationContext.Split("|".ToCharArray());

                    appDomainInfo.ApplicationBase     = s_ActivationContext[constBaseDir];
                    appDomainInfo.ApplicationName     = s_ActivationContext[constAppName];
                    appDomainInfo.ConfigurationFile   = s_ActivationContext[constCfgFile];
                    appDomainInfo.PrivateBinPath      = s_ActivationContext[constBinPath];
                    appDomainInfo.TargetFrameworkName = s_ActivationContext[constTgtFramework];
                    if (string.IsNullOrEmpty(appDomainInfo.TargetFrameworkName))
                    {
                        var asm = this.EntryAssembly;
                        if (asm != null)
                        {
                            TargetFrameworkAttribute[] attrs = (TargetFrameworkAttribute[])asm.GetCustomAttributes(typeof(TargetFrameworkAttribute));
                            if (attrs != null && attrs.Length > 0)
                            {
                                appDomainInfo.TargetFrameworkName = attrs[0].FrameworkName;
                            }
                        }
                    }
                    if (string.IsNullOrEmpty(appDomainInfo.TargetFrameworkName))
                    {
                        appDomainInfo.TargetFrameworkName = $".NETFramework,Version=v{NetVersion()}";
                    }
                }
            }

            // Per .Net documentation, if PrivateBinPath is not a subfolder of the ApplicationBase path, it is ignored.
            // To support arbitrary configured paths (my current default is the application base to be a root script folder and bin path of whatever specified)
            // the assembly resolve code is needed.
            // As an added benefit, it solves bindings issue since the assembly resolution code ignores version information and solely based on
            // the dll or exe file that can be resolved based on resolution search logic.
            // NOTE: If the developer add path to binaries to PATH variable, .Net loads an assembly prior of calling AssemblyResolve (and that code may 
            // still require bindings), but if that was the case this class will not be invoked (since it removes the need of adding bin location to PATH variable).
            // NOTE: this is needed for all application domains created by the process - including added by .Net code.
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

            base.InitializeNewDomain(appDomainInfo);
        }

        #region Resolving .Net version

        // despite looking hacky, it is actually recommended way:
        // https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed
        private string NetVersion()
        {
            // Checking the version using >= enables forward compatibility.
            string CheckFor45PlusVersion(int releaseKey)
            {
                if (releaseKey >= 528040)
                    return "4.8";
                if (releaseKey >= 461808)
                    return "4.7.2";
                if (releaseKey >= 461308)
                    return "4.7.1";
                if (releaseKey >= 460798)
                    return "4.7";
                if (releaseKey >= 394802)
                    return "4.6.2";
                if (releaseKey >= 394254)
                    return "4.6.1";
                if (releaseKey >= 393295)
                    return "4.6";
                if (releaseKey >= 379893)
                    return "4.5.2";
                if (releaseKey >= 378675)
                    return "4.5.1";
                if (releaseKey >= 378389)
                    return "4.5";

                // This code should never execute. A non-null release key should mean
                // that 4.5 or later is installed.
                return "4.0";
            }
            const string subkey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\";
            using (var ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(subkey))
            {
                if (ndpKey != null && ndpKey.GetValue("Release") != null)
                {
                    return CheckFor45PlusVersion((int)ndpKey.GetValue("Release"));
                }
                else
                {
                    return "4.0";
                }
            }
        }

        #endregion

        #region Resolving assembly using locations priority and bin path outside of ApplicationBase

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblyFullName = args.Name;

            // Ignore missing resources
            if (assemblyFullName.Contains(".resources"))
                return null;

            // check for assemblies already loaded
            Assembly resolvedAssembly = null;
            var asms = AppDomain.CurrentDomain.GetAssemblies();
            foreach (var asm in asms)
            {
                if (asm.FullName == assemblyFullName)
                {
                    resolvedAssembly = asm;
                    break;
                }
            }
            if (resolvedAssembly != null)
                return resolvedAssembly;

            // Try to load by filename - split out the filename of the full assembly name
            // and append the base path of the original assembly (ie. look in the same dir)
            string assemblyExpectedFileName = assemblyFullName.Split(',')[0];

            Assembly AssemblyResolverInFolders()
            {
                string AssemblyLocation(Assembly asm)
                {
                    if (asm != null)
                    {
                        if (asm.IsDynamic)
                            return "";
                        try
                        {
                            var location = asm.Location;
                            if (!string.IsNullOrEmpty(location))
                                return System.IO.Path.GetDirectoryName(location);
                        }
                        catch { }
                    }
                    return "";
                }

                // search folders priority:
                // 1) AppBase (we expect it to have):
                //    kind of default, kind of redundant - since .Net weould resolve it for us any way;
                //    but if assembly was rejected due to version mismatch, our file-based code would pick it up
                // 2) PrivateBinPath defined for the AppDomain (if defined):
                //    standard .Net rules consider this location only if inside ApplicationBase; we use it always
                // 3) Path of requesting assembly
                // 4) Current / working directory of the process (as it was defined when activating the default domain)
                // 5) Current / working directory at the time of resolution
                // 6) "Fair game folders" - locations of all loaded assemblies
                // 7) Subfolders of each folder included in the search folders priority
                // NOTE: The flexibility here is rather unlimited - the activation context can pass unlimited number of switches
                //       (at least 64 if we want to limit ourselves to only a single activation context token)
                List<string> lstSearchFolders = new List<string>();
                if (!string.IsNullOrEmpty(AppDomain.CurrentDomain.BaseDirectory))
                    lstSearchFolders.Add(AppDomain.CurrentDomain.BaseDirectory);
                if (!string.IsNullOrEmpty(AppDomain.CurrentDomain.SetupInformation.PrivateBinPath))
                    lstSearchFolders.Add(AppDomain.CurrentDomain.SetupInformation.PrivateBinPath);

                var locationAsm = AssemblyLocation(args.RequestingAssembly);
                if (!string.IsNullOrEmpty(locationAsm))
                {
                    lstSearchFolders.Add(locationAsm);
                }

                lstSearchFolders.Add(s_ActivationContext[constWorkingDir]);
                lstSearchFolders.Add(System.Environment.CurrentDirectory);
                
                foreach (var a in AppDomain.CurrentDomain.GetAssemblies())
                {
                    locationAsm = AssemblyLocation(a);
                    if (!string.IsNullOrEmpty(locationAsm))
                    {
                        lstSearchFolders.Add(locationAsm);
                    }
                }

                // remove duplicates
                lstSearchFolders = lstSearchFolders.Distinct().ToList();

                // add subfolders for each search folder
                List<string> lstSubFolders = new List<string>();
                void CollectDirectoryRecursive(List<string> collection, string directory, bool includeItself = true)
                {
                    // add itself
                    if (includeItself)
                        collection.Add(directory);

                    // recurse
                    try
                    {
                        var dir = new System.IO.DirectoryInfo(directory);
                        var subfolders = dir.GetDirectories();
                        foreach (var di in subfolders)
                        {
                            CollectDirectoryRecursive(collection, di.FullName);
                        }
                    }
                    catch { }
                }
                foreach (var mainFolder in lstSearchFolders)
                {
                    CollectDirectoryRecursive(lstSubFolders, mainFolder, false);
                }
                lstSearchFolders.AddRange(lstSubFolders);

                // in the order of provided folders, locate expected file name for the assembly
                // load using the first find
                foreach (var folder in lstSearchFolders)
                {
                    var fullAsmFileName = System.IO.Path.Combine(folder, assemblyExpectedFileName);
                    string[] assemblies = new string[] {
                        fullAsmFileName + ".dll",
                        fullAsmFileName + ".exe"
                    };

                    foreach (var asmFile in assemblies)
                    {
                        if (System.IO.File.Exists(asmFile))
                        {
                            try
                            {
                                var assembly = Assembly.LoadFrom(asmFile);
                                // using the trace or debug facility may trigger circular dependencies in case the trace listener is redirected to one managed by the code
                                // System.Diagnostics.Debug.WriteLine($"Assembly Resolver: {assemblyExpectedFileName} found in: {fullAsmFileName} Loaded as {assembly.Location}");
                                return assembly;
                            }
                            catch { }
                        }
                    }
                }
                return null;
            }
            return AssemblyResolverInFolders();
        }

        #endregion
    }
}
