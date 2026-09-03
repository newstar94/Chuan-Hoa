using System;
using System.ComponentModel;
using System.Diagnostics;
using Microsoft.Office.Tools;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Word = Microsoft.Office.Interop.Word;

namespace ChuanHoa.AddIn.Vsto
{
    [StartupObject(0)]
    public partial class ThisAddIn : AddInBase
    {
        [DebuggerNonUserCode]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public ThisAddIn(Factory factory, IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Word.Application Application
        {
            get { return GetHostItem<Word.Application>(typeof(Word.Application), "Application"); }
        }

        [DebuggerNonUserCode]
        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void Initialize()
        {
            base.Initialize();
        }

        [DebuggerNonUserCode]
        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void FinishInitialization()
        {
            InternalStartup();
            OnStartup();
        }

        [DebuggerNonUserCode]
        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void InitializeDataBindings()
        {
        }

        [DebuggerNonUserCode]
        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void OnShutdown()
        {
            base.OnShutdown();
        }
    }
}
