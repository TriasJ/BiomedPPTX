using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using System.Reflection;

using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for THandler
    /// </summary>
    /// <typeparam name="THandler">Target handler type</typeparam>
    public abstract class BaseHandlerFactory<THandler>
    {
        [ImportMany]
        private IEnumerable<Lazy<THandler, IRibbonIdMetadata>> ImportedHandlers { get; set; }

        protected BaseHandlerFactory()
        {
            try
            {
                AggregateCatalog catalog = new AggregateCatalog(
                    new AssemblyCatalog(Assembly.GetExecutingAssembly()));
                CompositionContainer container = new CompositionContainer(catalog);
                container.ComposeParts(this);
            }
            catch (ReflectionTypeLoadException ex)
            {
                string loaderMessages = "";
                if (ex.LoaderExceptions != null)
                {
                    foreach (Exception le in ex.LoaderExceptions)
                    {
                        if (le != null)
                        {
                            loaderMessages += le.Message + "\n";
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine("MEF ReflectionTypeLoadException: " + loaderMessages);
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "BiomedPPTX_MEF_Error.txt"),
                    "ReflectionTypeLoadException:\n" + loaderMessages + "\n\nStackTrace:\n" + ex.StackTrace);

                AggregateCatalog safeCatalog = new AggregateCatalog();
                foreach (Type type in GetLoadableTypes(Assembly.GetExecutingAssembly()))
                {
                    try
                    {
                        safeCatalog.Catalogs.Add(new TypeCatalog(type));
                    }
                    catch
                    {
                    }
                }

                CompositionContainer container = new CompositionContainer(safeCatalog);
                container.ComposeParts(this);
            }
        }

        public THandler CreateInstance(string ribbonId, string ribbonTag)
        {
            foreach (Lazy<THandler, IRibbonIdMetadata> handler in ImportedHandlers)
            {
                if (handler.Metadata.RibbonIds.Contains(ribbonId)
                    || handler.Metadata.RibbonIds.Contains(ribbonTag))
                {
                    return handler.Value;
                }
            }

            return GetEmptyHandler();
        }

        protected abstract THandler GetEmptyHandler();

        private static IEnumerable<Type> GetLoadableTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(t => t != null);
            }
        }
    }
}
