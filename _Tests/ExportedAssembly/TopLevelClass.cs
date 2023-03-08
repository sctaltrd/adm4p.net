using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using NXPorts.Attributes;

namespace ExportedAssembly
{
    public class TopLevelClass
    {
        [DllExport("tlc_BringItUp", CallingConvention.Cdecl)]
        public static void BringItUp()
        {
            try
            {
                int[] test = new int[0x3FFFFFFF];
            }
            catch { }

            var ad = AppDomain.CurrentDomain;
            var actfm = AppContext.TargetFrameworkName;

            // this is the test if default domain is properly initialized:
            // this compatibility switch is set for true with target framework 4.7.2 and older and should be false if 4.8 or newer is specified
            bool sw = false;
            AppContext.TryGetSwitch("Switch.System.Threading.ThrowExceptionIfDisposedCancellationTokenSource", out sw);

            // test of using the configuration system
            var testValue = ConfigurationManager.AppSettings["A"];

            // test for entry assembly (mscorlib is ecpected)
            var asm = Assembly.GetEntryAssembly();

            Debug.WriteLine($"BringItUop {DateTime.Now}");
        }
    }
}
