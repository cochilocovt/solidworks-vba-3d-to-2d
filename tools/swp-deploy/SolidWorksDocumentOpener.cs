using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace TargetSpecHybrid.SwpDeployment
{
    public sealed class OpenDocResult
    {
        public bool Succeeded { get; set; }
        public int Errors { get; set; }
        public int Warnings { get; set; }
        public string Title { get; set; }
    }

    public sealed class OpenDocumentInfo
    {
        public string Title { get; set; }
        public string PathName { get; set; }
    }

    // PowerShell's native COM late binding fails OpenDoc6 with
    // TYPE_E_ELEMENTNOTFOUND (its out-parameter resolution cannot match
    // the overload), and a direct PowerShell bracket-cast of the raw
    // GetActiveObject() proxy to ISldWorks also fails even after the
    // interop assembly is loaded via Add-Type -- a live-run finding
    // recorded in docs/SOLIDWORKS_API_VALIDATION.md. Routing the call
    // through compiled C#, exactly like SolidWorksMacroInvoker.Run does
    // for RunMacro2, is what actually works.
    public static class SolidWorksDocumentOpener
    {
        public static OpenDocumentInfo[] ListOpenDocuments(object application)
        {
            try
            {
                return ListOpenDocumentsEarlyBound(application);
            }
            catch (InvalidCastException)
            {
                return ListOpenDocumentsDispatch(application);
            }
            catch (COMException exception)
            {
                const int NoInterface = unchecked((int)0x80004002);
                if (exception.ErrorCode != NoInterface)
                {
                    throw;
                }

                return ListOpenDocumentsDispatch(application);
            }
        }

        public static OpenDocResult Open(
            object application,
            string fileName,
            int documentType,
            int options)
        {
            try
            {
                return OpenEarlyBound(
                    application,
                    fileName,
                    documentType,
                    options);
            }
            catch (InvalidCastException)
            {
                return OpenDispatch(
                    application,
                    fileName,
                    documentType,
                    options);
            }
            catch (COMException exception)
            {
                const int NoInterface = unchecked((int)0x80004002);
                if (exception.ErrorCode != NoInterface)
                {
                    throw;
                }

                return OpenDispatch(
                    application,
                    fileName,
                    documentType,
                    options);
            }
        }

        private static OpenDocumentInfo[] ListOpenDocumentsEarlyBound(
            object application)
        {
            ISldWorks solidWorks = (ISldWorks)application;
            var documents = new System.Collections.Generic.List<
                OpenDocumentInfo>();
            ModelDoc2 document = solidWorks.GetFirstDocument() as ModelDoc2;

            while (document != null)
            {
                documents.Add(new OpenDocumentInfo
                {
                    Title = document.GetTitle(),
                    PathName = document.GetPathName()
                });
                document = document.GetNext() as ModelDoc2;
            }

            return documents.ToArray();
        }

        private static OpenDocumentInfo[] ListOpenDocumentsDispatch(
            object application)
        {
            var documents = new System.Collections.Generic.List<
                OpenDocumentInfo>();
            object document = application.GetType().InvokeMember(
                "GetFirstDocument",
                BindingFlags.InvokeMethod,
                null,
                application,
                new object[0],
                CultureInfo.InvariantCulture);

            while (document != null)
            {
                string title = Convert.ToString(
                    document.GetType().InvokeMember(
                        "GetTitle",
                        BindingFlags.InvokeMethod,
                        null,
                        document,
                        new object[0],
                        CultureInfo.InvariantCulture),
                    CultureInfo.InvariantCulture);
                string pathName = Convert.ToString(
                    document.GetType().InvokeMember(
                        "GetPathName",
                        BindingFlags.InvokeMethod,
                        null,
                        document,
                        new object[0],
                        CultureInfo.InvariantCulture),
                    CultureInfo.InvariantCulture);
                documents.Add(new OpenDocumentInfo
                {
                    Title = title,
                    PathName = pathName
                });
                document = document.GetType().InvokeMember(
                    "GetNext",
                    BindingFlags.InvokeMethod,
                    null,
                    document,
                    new object[0],
                    CultureInfo.InvariantCulture);
            }

            return documents.ToArray();
        }

        private static OpenDocResult OpenEarlyBound(
            object application,
            string fileName,
            int documentType,
            int options)
        {
            ISldWorks solidWorks = (ISldWorks)application;
            int errors = 0;
            int warnings = 0;

            ModelDoc2 document = solidWorks.OpenDoc6(
                fileName,
                documentType,
                options,
                string.Empty,
                ref errors,
                ref warnings);

            if (document != null)
            {
                int activationErrors = 0;
                solidWorks.ActivateDoc3(
                    document.GetTitle(),
                    false,
                    0,
                    ref activationErrors);
            }

            return new OpenDocResult
            {
                Succeeded = document != null,
                Errors = errors,
                Warnings = warnings,
                Title = document != null ? document.GetTitle() : null
            };
        }

        private static OpenDocResult OpenDispatch(
            object application,
            string fileName,
            int documentType,
            int options)
        {
            object[] arguments =
            {
                fileName,
                documentType,
                options,
                string.Empty,
                0,
                0
            };

            ParameterModifier modifier = new ParameterModifier(arguments.Length);
            modifier[4] = true;
            modifier[5] = true;

            object rawResult = application.GetType().InvokeMember(
                "OpenDoc6",
                BindingFlags.InvokeMethod,
                null,
                application,
                arguments,
                new[] { modifier },
                CultureInfo.InvariantCulture,
                null);

            string title = null;
            if (rawResult != null)
            {
                title = Convert.ToString(
                    rawResult.GetType().InvokeMember(
                        "GetTitle",
                        BindingFlags.InvokeMethod,
                        null,
                        rawResult,
                        new object[0],
                        CultureInfo.InvariantCulture),
                    CultureInfo.InvariantCulture);

                object[] activateArguments = { title, false, 0, 0 };
                ParameterModifier activateModifier = new ParameterModifier(
                    activateArguments.Length);
                activateModifier[3] = true;
                application.GetType().InvokeMember(
                    "ActivateDoc3",
                    BindingFlags.InvokeMethod,
                    null,
                    application,
                    activateArguments,
                    new[] { activateModifier },
                    CultureInfo.InvariantCulture,
                    null);
            }

            return new OpenDocResult
            {
                Succeeded = rawResult != null,
                Errors = Convert.ToInt32(
                    arguments[4],
                    CultureInfo.InvariantCulture),
                Warnings = Convert.ToInt32(
                    arguments[5],
                    CultureInfo.InvariantCulture),
                Title = title
            };
        }
    }
}
