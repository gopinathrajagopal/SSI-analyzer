namespace analyzer.Properties
{
    using System;
    using System.CodeDom.Compiler;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.Resources;
    using System.Runtime.CompilerServices;

    [CompilerGenerated, DebuggerNonUserCode, GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    internal class Resources
    {
        private static CultureInfo resourceCulture;
        private static System.Resources.ResourceManager resourceMan;

        internal Resources()
        {
        }

        internal static Icon analyzer
        {
            get
            {
                return (Icon) ResourceManager.GetObject("analyzer", resourceCulture);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        internal static CultureInfo Culture
        {
            get
            {
                return resourceCulture;
            }
            set
            {
                resourceCulture = value;
            }
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        internal static System.Resources.ResourceManager ResourceManager
        {
            get
            {
                if (object.ReferenceEquals(resourceMan, null))
                {
                    System.Resources.ResourceManager manager = new System.Resources.ResourceManager("analyzer.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = manager;
                }
                return resourceMan;
            }
        }

        internal static Bitmap vizOff
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("vizOff", resourceCulture);
            }
        }

        internal static Bitmap vizOn
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("vizOn", resourceCulture);
            }
        }

        internal static Icon winConnect
        {
            get
            {
                return (Icon) ResourceManager.GetObject("winConnect", resourceCulture);
            }
        }

        internal static Icon winDiscon
        {
            get
            {
                return (Icon) ResourceManager.GetObject("winDiscon", resourceCulture);
            }
        }
    }
}

