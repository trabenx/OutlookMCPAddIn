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
                // Access the Outlook Application object.
                // In VSTO, 'Globals.ThisAddIn.Application' is a reliable way to get it.
                // Or, if your VSTO project directly exposes it on 'this', then 'this.Application' is fine.
                // Since 'this.Application' gave an error, let's try Globals.
                Outlook.Application outlookApplication = this.Application; // If this.Application is truly missing, this line is the problem source
                                                                           // Let's assume for a moment the issue is with the event wiring, not Application access.
                _syncContext = SynchronizationContext.Current;
                if (_outlookController == null)
                {
                    // Pass the application object obtained above
                    _outlookController = new OutlookController(outlookApplication);
                }

                McpHttpServer.Start(_outlookController, _syncContext);
                System.Diagnostics.Debug.WriteLine("OutlookMcpAddIn (VSTO) Started and MCP Server Initialized.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ThisAddIn_Startup: {ex.ToString()}");
                // Consider user notification for critical startup failures
                // System.Windows.Forms.MessageBox.Show($"Failed to start MCP Add-in: {ex.Message}", "Add-in Error");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                McpHttpServer.Stop();
                _outlookController = null;
                _syncContext = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Diagnostics.Debug.WriteLine("OutlookMcpAddIn (VSTO) Shutdown initiated and MCP Server Stopped.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ThisAddIn_Shutdown: {ex.ToString()}");
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            // These lines are crucial. If 'Startup' and 'Shutdown' events are not found on 'this',
            // it means the 'ThisAddIn' partial class definition is missing them or they are
            // defined in a way the compiler isn't recognizing from this context.
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
