using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace TargetSpecHybrid.SwpDeployment
{
    public sealed class MacroRunResult
    {
        public bool Succeeded { get; set; }
        public int Error { get; set; }
    }

    public static class SolidWorksMacroInvoker
    {
        public static MacroRunResult Run(
            object application,
            string macroPath,
            string moduleName,
            string procedureName,
            int options)
        {
            try
            {
                return RunEarlyBound(
                    application,
                    macroPath,
                    moduleName,
                    procedureName,
                    options);
            }
            catch (InvalidCastException)
            {
                return RunDispatch(
                    application,
                    macroPath,
                    moduleName,
                    procedureName,
                    options);
            }
            catch (COMException exception)
            {
                const int NoInterface = unchecked((int)0x80004002);
                if (exception.ErrorCode != NoInterface)
                {
                    throw;
                }

                return RunDispatch(
                    application,
                    macroPath,
                    moduleName,
                    procedureName,
                    options);
            }
        }

        private static MacroRunResult RunEarlyBound(
            object application,
            string macroPath,
            string moduleName,
            string procedureName,
            int options)
        {
            ISldWorks solidWorks = (ISldWorks)application;
            int error = 0;
            bool succeeded = solidWorks.RunMacro2(
                macroPath,
                moduleName,
                procedureName,
                options,
                out error);

            return new MacroRunResult
            {
                Succeeded = succeeded,
                Error = error
            };
        }

        private static MacroRunResult RunDispatch(
            object application,
            string macroPath,
            string moduleName,
            string procedureName,
            int options)
        {
            object[] arguments =
            {
                macroPath,
                moduleName,
                procedureName,
                options,
                0
            };

            ParameterModifier modifier = new ParameterModifier(arguments.Length);
            modifier[4] = true;

            object rawResult = application.GetType().InvokeMember(
                "RunMacro2",
                BindingFlags.InvokeMethod,
                null,
                application,
                arguments,
                new[] { modifier },
                CultureInfo.InvariantCulture,
                null);

            return new MacroRunResult
            {
                Succeeded = Convert.ToBoolean(
                    rawResult,
                    CultureInfo.InvariantCulture),
                Error = Convert.ToInt32(
                    arguments[4],
                    CultureInfo.InvariantCulture)
            };
        }
    }
}
