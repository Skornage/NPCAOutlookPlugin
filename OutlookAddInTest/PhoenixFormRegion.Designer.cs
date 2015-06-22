namespace OutlookAddInTest
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class PhoenixFormRegion : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public PhoenixFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // FormRegion1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Name = "FormRegion1";
            this.Size = new System.Drawing.Size(669, 45);
            this.FormRegionShowing += new System.EventHandler(this.FormRegion1_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.FormRegion1_FormRegionClosed);
            this.ResumeLayout(false);

        }

        #endregion

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PhoenixFormRegion));
            manifest.ExactMessageClass = true;
            manifest.FormRegionName = "FormRegion1";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Replacement;
            manifest.Icons.Default = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Default")));
            manifest.Icons.Encrypted = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Encrypted")));
            manifest.Icons.Forwarded = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Forwarded")));
            manifest.Icons.Page = global::OutlookAddInTest.Properties.Resources.apple_touch_icon;
            manifest.Icons.Read = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Read")));
            manifest.Icons.Recurring = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Recurring")));
            manifest.Icons.Replied = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Replied")));
            manifest.Icons.Signed = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Signed")));
            manifest.Icons.Submitted = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Submitted")));
            manifest.Icons.Unread = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Unread")));
            manifest.Icons.Unsent = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Unsent")));
            manifest.Icons.Window = ((System.Drawing.Icon)(resources.GetObject("FormRegion1.Manifest.Icons.Window")));
            manifest.ShowInspectorCompose = false;
            manifest.ShowInspectorRead = false;
            manifest.Title = "FormRegion1";

        }

        #endregion

        public partial class FormRegion1Factory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public FormRegion1Factory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                PhoenixFormRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.FormRegion1Factory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                PhoenixFormRegion form = new PhoenixFormRegion(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal PhoenixFormRegion FormRegion1
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(PhoenixFormRegion))
                        return (PhoenixFormRegion)item;
                }
                return null;
            }
        }
    }
}
