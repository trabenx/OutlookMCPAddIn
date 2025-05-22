using OutlookMCPAddIn;
using System;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
// using Office = Microsoft.Office.Core; // Usually not needed unless using CommandBars etc.

namespace OutlookMcpAddIn // Ensure this namespace matches your project
{
    public partial class ThisAddIn
    {
        private static OutlookController _outlookController;
        private static SynchronizationContext _outlookMainThreadSyncContext; // This will be our reliable one
        private Control _hiddenControlForContext; // Helper to get context

        // The ThisAddIn_Startup and ThisAddIn_Shutdown methods are event handlers.
        // They are wired up in the InternalStartup method, which IS CALLED by VSTO.

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Outlook.Application outlookApplication = Globals.ThisAddIn.Application;

                // Create a hidden control on the current thread (which is Outlook's main thread for Startup)
                // This control's Invoke/BeginInvoke will marshal to this thread.
                _hiddenControlForContext = new Control();
                _hiddenControlForContext.CreateControl(); // Ensures the handle is created

                // Get the SynchronizationContext from this control
                // However, a Control itself is an ISynchronizeInvoke provider.
                // It's often more direct to pass the control itself for marshalling.

                // Let's try getting SynchronizationContext from the control's thread
                // This is a bit indirect. A better way is to use the control directly for Invoke.
                // SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext()); // Don't do this globally
                // _outlookMainThreadSyncContext = SynchronizationContext.Current; // Capture after creating control

                // Simpler: Pass the control to McpHttpServer and let it use control.Invoke/BeginInvoke
                // Or, if McpHttpServer strictly needs a SynchronizationContext object:
                if (_hiddenControlForContext.IsHandleCreated)
                {
                    // This is a way to get a WindowsFormsSynchronizationContext if one isn't current
                    // but it's usually better to use the control's ISynchronizeInvoke methods.
                    // Forcing one like this might not be ideal.
                    // Let's try to capture SynchronizationContext.Current AGAIN after creating the control.
                    // Often, creating a WinForms control on a thread installs a WindowsFormsSynchronizationContext if one isn't there.
                    _outlookMainThreadSyncContext = SynchronizationContext.Current;
                }


                if (_outlookMainThreadSyncContext == null)
                {
                    // Fallback: If SynchronizationContext.Current is STILL null even after creating a control,
                    // this is highly unusual for the thread running ThisAddIn_Startup.
                    // We might have to resort to a timer polling mechanism for McpHttpServer.
                    System.Diagnostics.Debug.WriteLine("CRITICAL: Still could not obtain a SynchronizationContext even after creating a hidden control. MCP Server will not start with marshalling.");
                    // For now, prevent server start if this fails, as it's fundamental.
                    return;
                }


                if (_outlookController == null)
                {
                    _outlookController = new OutlookController(outlookApplication);
                }

                // Pass the captured (hopefully valid) _outlookMainThreadSyncContext
                McpHttpServer.Start(_outlookController, _outlookMainThreadSyncContext);
                System.Diagnostics.Debug.WriteLine("OutlookMcpAddIn (VSTO) Started and MCP Server Initialized with context from hidden control.");

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ThisAddIn_Startup: {ex.ToString()}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                McpHttpServer.Stop();
                if (_hiddenControlForContext != null)
                {
                    _hiddenControlForContext.Dispose();
                    _hiddenControlForContext = null;
                }
                _outlookController = null;
                _outlookMainThreadSyncContext = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Diagnostics.Debug.WriteLine("OutlookMcpAddIn (VSTO) Shutdown initiated.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ThisAddIn_Shutdown: {ex.ToString()}");
            }
        }


        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
