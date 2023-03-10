using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.IO;
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

        const string constActivationContextEnvVariable = "ADM_4_PYTHON_ACTIVATION_CONTEXT";

        const int constWorkingDir = 0;
        const int constBaseDir = 1;
        const int constAppName = 2;
        const int constBinPath = 3;
        const int constCfgFile = 4;
        const int constTgtFramework = 5;
        const int constCulture = 6;
        const int constSwitchesAR = 7;
        const int constSwitchesGeneral = 8;

        private static string[] s_ActivationContext;

        const ulong SWITCH_AR_CONSIDERREQUESTINGASM = 0x0000000000000001;
        const ulong SWITCH_AR_WORKINGDIR            = 0x0000000000000002;
        const ulong SWITCH_AR_ACTIVATIONWORKINGDIR  = 0x0000000000000004;
        const ulong SWITCH_AR_LOADEDASMS            = 0x0000000000000008;
        const ulong SWITCH_AR_SUBFOLDERS            = 0x0000000000000010;
        private static ulong s_SwitchesAR           = 0xFFFFFFFFFFFFFFFF;

        private static ulong s_SwitchesGeneral = 0;

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
                string activationContext = System.Environment.GetEnvironmentVariable(constActivationContextEnvVariable);
                if (!string.IsNullOrEmpty(activationContext))
                {
                    s_AsmMscorlib = typeof(Object).Assembly;
                    s_ActivationContext = activationContext.Split("|".ToCharArray());

                    // mandatory tokens
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

                    // culture
                    string culture = s_ActivationContext[constCulture];
                    if (!string.IsNullOrEmpty(culture))
                    {
                        try
                        {
                            var cultureInfo = CultureInfo.GetCultureInfo(culture);
                            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
                        }
                        catch { }
                    }

                    // switches
                    string switchesAR = s_ActivationContext[constSwitchesAR];
                    if (!string.IsNullOrEmpty(switchesAR))
                    {
                        try
                        {
                            s_SwitchesAR = ulong.Parse(switchesAR);
                        }
                        catch { }
                    }
                    string switchesGeneral = s_ActivationContext[constSwitchesGeneral];
                    if (!string.IsNullOrEmpty(switchesGeneral))
                    {
                        try
                        {
                            s_SwitchesGeneral = ulong.Parse(switchesGeneral);
                        }
                        catch { }
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
            var ad = AppDomain.CurrentDomain;
            var assemblyFullName = args.Name;

            // Ignore missing resources
            if (assemblyFullName.Contains(".resources"))
                return null;

            // check for assemblies already loaded
            Assembly resolvedAssembly = null;
            var asms = ad.GetAssemblies();
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

            string AssemblyResolutionCultureName()
            {
                // check for culture of the requested assembly
                var aName = new AssemblyName(assemblyFullName);
                var ci = aName.CultureInfo;
                if (ci != null && !ci.IsNeutralCulture)
                    return ci.Name;

                // check for culture of the current thread
                ci = CultureInfo.CurrentCulture;
                if (ci != null && !ci.IsNeutralCulture)
                    return ci.Name;

                // check for culture of the requesting assembly
                if (args.RequestingAssembly != null && (s_SwitchesAR & SWITCH_AR_CONSIDERREQUESTINGASM) > 0)
                {
                    try
                    {
                        ci = args.RequestingAssembly.GetName().CultureInfo;
                        if (ci != null && !ci.IsNeutralCulture)
                            return ci.Name;
                    }
                    catch { }
                }

                return null;
            }

            string cultureName = AssemblyResolutionCultureName(); 

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
                                return Path.GetDirectoryName(location);
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
                // NOTE: code handles culture specific assemblies if requested

                // NOTE: while code is generally adheres to standard .Net rules per
                //       https://learn.microsoft.com/en-us/dotnet/framework/deployment/how-the-runtime-locates-assemblies,
                //       we also extend search logic to accomodate a hybrid nature of CLR hosted by Python scripting environment

                List<string> lstSearchFolders = new List<string>();

                // Primary folders.
                // Order of priority is pre-set and they are not optional
                if (!string.IsNullOrEmpty(ad.BaseDirectory))
                {
                    lstSearchFolders.Add(ad.BaseDirectory);
                    if (cultureName != null)
                    {
                        lstSearchFolders.Add(Path.Combine(ad.BaseDirectory, cultureName));
                    }
                }
                if (!string.IsNullOrEmpty(ad.SetupInformation.PrivateBinPath))
                {
                    lstSearchFolders.Add(ad.SetupInformation.PrivateBinPath);
                    if (cultureName != null)
                    {
                        lstSearchFolders.Add(Path.Combine(ad.SetupInformation.PrivateBinPath, cultureName));
                    }
                }

                // Location of requesting assembly - switchable
                if (args.RequestingAssembly != null && (s_SwitchesAR & SWITCH_AR_CONSIDERREQUESTINGASM) > 0)
                {
                    var locationAsm = AssemblyLocation(args.RequestingAssembly);
                    if (!string.IsNullOrEmpty(locationAsm))
                    {
                        lstSearchFolders.Add(locationAsm);
                        if (cultureName != null)
                        {
                            lstSearchFolders.Add(Path.Combine(locationAsm, cultureName));
                        }
                    }
                }

                // Current directory - current working dir (normal OS stuff) as well as working dir at the time of AppDomain loading
                if ((s_SwitchesAR & SWITCH_AR_ACTIVATIONWORKINGDIR) > 0)
                {
                    lstSearchFolders.Add(s_ActivationContext[constWorkingDir]);
                    if (cultureName != null)
                    {
                        lstSearchFolders.Add(Path.Combine(s_ActivationContext[constWorkingDir], cultureName));
                    }
                }
                if ((s_SwitchesAR & SWITCH_AR_WORKINGDIR) > 0)
                {
                    lstSearchFolders.Add(Environment.CurrentDirectory);
                    if (cultureName != null)
                    {
                        lstSearchFolders.Add(Path.Combine(Environment.CurrentDirectory, cultureName));
                    }
                }

                // Locations of currently loaded assemblies - switchable
                if ((s_SwitchesAR & SWITCH_AR_LOADEDASMS) > 0)
                {
                    foreach (var a in ad.GetAssemblies())
                    {
                        var locationAsm = AssemblyLocation(a);
                        if (!string.IsNullOrEmpty(locationAsm))
                        {
                            lstSearchFolders.Add(locationAsm);
                        }
                    }
                }

                // remove duplicates
                lstSearchFolders = lstSearchFolders.Distinct().ToList();

                // add subfolders for each search folder - switchable
                if ((s_SwitchesAR & SWITCH_AR_SUBFOLDERS) > 0)
                {
                    List<string> lstSubFolders = new List<string>();
                    void CollectDirectoryRecursive(List<string> collection, string directory, bool includeItself = true)
                    {
                        // add itself
                        if (includeItself)
                            collection.Add(directory);

                        // recurse
                        try
                        {
                            var dir = new DirectoryInfo(directory);
                            if (Directory.Exists(dir.FullName))
                            {
                                var subfolders = dir.GetDirectories();
                                foreach (var di in subfolders)
                                {
                                    CollectDirectoryRecursive(collection, di.FullName);
                                }
                            }
                        }
                        catch { }
                    }
                    foreach (var mainFolder in lstSearchFolders)
                    {
                        CollectDirectoryRecursive(lstSubFolders, mainFolder, false);
                    }
                    lstSearchFolders.AddRange(lstSubFolders);
                }

                // final cleanup of duplicates
                lstSearchFolders = lstSearchFolders.Distinct().ToList();

                // in the order of provided folders, locate expected file name for the assembly
                // load using the first find
                foreach (var folder in lstSearchFolders)
                {
                    var fullAsmFileName = Path.Combine(folder, assemblyExpectedFileName);
                    string[] assemblies = new string[] {
                        fullAsmFileName,
                        fullAsmFileName + ".dll",
                        fullAsmFileName + ".exe"
                    };

                    foreach (var asmFile in assemblies)
                    {
                        if (File.Exists(asmFile))
                        {
                            try
                            {
                                var assembly = Assembly.LoadFrom(asmFile);
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
