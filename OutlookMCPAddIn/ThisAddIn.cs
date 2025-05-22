using OutlookMCPAddIn;
using System;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
// using Office = Microsoft.Office.Core; // Usually not needed unless using CommandBars etc.

namespace OutlookMcpAddIn // Ensure this namespace matches your project
{
    public partial class ThisAddIn
    {
        private static OutlookController _outlookController;
        private static SynchronizationContext _syncContext;

        // The ThisAddIn_Startup and ThisAddIn_Shutdown methods are event handlers.
        // They are wired up in the InternalStartup method, which IS CALLED by VSTO.

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Outlook.Application outlookApplication = Globals.ThisAddIn.Application;
                _syncContext = SynchronizationContext.Current; // Attempt to capture

                if (_outlookController == null)
                {
                    _outlookController = new OutlookController(outlookApplication);
                }

                if (_syncContext != null)
                {
                    // Only attempt to start the server if we have a sync context.
                    // McpHttpServer.Start still needs to be robust if _syncContext is null,
                    // OR McpHttpServer.Start should throw if syncContext is null and it can't operate without it.
                    // For now, assuming McpHttpServer.Start will handle a null syncContext argument (as modified below).
                    McpHttpServer.Start(_outlookController, _syncContext);
                    System.Diagnostics.Debug.WriteLine("OutlookMcpAddIn (VSTO) Started and MCP Server Initialized.");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("OutlookMcpAddIn (VSTO) Started, but MCP Server NOT Initialized: SynchronizationContext.Current was null. Calls to Outlook may fail.");
                    // Optionally, inform the user or log this as a critical failure for the MCP functionality.
                    // For example, you could display a MessageBox:
                    // System.Windows.Forms.MessageBox.Show("A required component (SynchronizationContext) for AI features could not be initialized. Some functionality may be impaired.", "Add-in Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ThisAddIn_Startup: {ex.ToString()}");
                // System.Windows.Forms.MessageBox.Show($"Critical error during add-in startup: {ex.Message}", "Add-in Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                McpHttpServer.Stop(); // McpHttpServer.Stop() should be safe even if Start wasn't fully successful
                _outlookController = null;
                _syncContext = null; // Clear it

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
