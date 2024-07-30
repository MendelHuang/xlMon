using System;

namespace XLMonCOMAddin
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    [System.Runtime.InteropServices.ProgId("HSBC_XLMonCOMAddin.MyConnect")]
    [System.Runtime.InteropServices.Guid("ffffffff-974d-44a3-8a5e-100000000001")]

    [System.Runtime.InteropServices.ComVisible(true)]
    public class XLMonCOMAddin : Extensibility.IDTExtensibility2
    {
        //Global variable for showing the execution sequence of each event
        public static int globalCounter = 0;
        public static string eventCallSequence = "";

        Microsoft.Office.Interop.Excel.Application AppObj;
        Microsoft.Office.Core.COMAddIn AddinInst;

        public void OnConnection(
            object Application,
            Extensibility.ext_ConnectMode ConnectMode,
            object AddInInst,
            ref Array custom)
        {
            try
            {
                this.AppObj = (Microsoft.Office.Interop.Excel.Application)Application;
                if (this.AddinInst == null)
                {
                    this.AddinInst = (Microsoft.Office.Core.COMAddIn)AddInInst;
                    this.AddinInst.Object = this;
                }

                //We are not relying on workbook close to calculate workbook open time as we would get no results if Excel.exe crashed or was terminated
                AppObj.SheetActivate += AppObj_SheetActivate;
                AppObj.AfterCalculate += AppObj_AfterCalculate;
                AppObj.WorkbookActivate += AppObj_WorkbookActivate;
                AppObj.SheetSelectionChange += AppObj_SheetSelectionChange;
                AppObj.WorkbookOpen += AppObj_WorkbookOpen;
            }
            catch (Exception ex)
            {
                // This is one of the few cases we log what we send over UDP differently to what we write to the file log.
                //Don't want to be sending exception string over UDP.
                System.Windows.Forms.MessageBox.Show("Exception found in OnConnection" + ex.ToString());
            }
        }

        public void OnDisconnection(
            Extensibility.ext_DisconnectMode RemoveMode,
            ref Array custom)
        {
            try
                {
                    AppObj.WorkbookOpen -= AppObj_WorkbookOpen;
                    AppObj.SheetSelectionChange -= AppObj_SheetSelectionChange;
                    AppObj.WorkbookActivate -= AppObj_WorkbookActivate;
                    AppObj.AfterCalculate -= AppObj_AfterCalculate;
                    AppObj.SheetActivate -= AppObj_SheetActivate;
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Exception found in OnDisconnection" + ex.ToString());
                }

                this.AppObj = null;
                this.AddinInst = null;
        }

        //Required Stubs so we are fully implementing the interface
        public void OnAddInsUpdate(ref Array custom) {}
        public void OnStartupComplete(ref Array custom) {}
        public void OnBeginShutdown(ref Array custom) {}

        private void AppObj_SheetActivate(object Sh)
        {
            //Show message box for this event
            IncrementAndDisplay("SheetActivate");
        }

        private void AppObj_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            //Show message box for this event
            IncrementAndDisplay("WorkbookOpen");
        }

        private void AppObj_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            //Show message box for this event
            IncrementAndDisplay("SheetSelectionChange");
        }

        private void AppObj_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            //Show message box for this event
            IncrementAndDisplay("WorkbookActivate");

            //This looks wrong, but the COM Addins implemented in C# seem to require us to do a memory clean up of any COM References we have manaually and it has to be called twice.
            //If we don't do this the workbooks stay visible in the VBA Editor and we are likely to get close down problems. This event seems like the best one to put it in.
            //We don't want this called too often.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void AppObj_AfterCalculate()
        {
            //Show message box for this event
            IncrementAndDisplay("AfterCalculate");
        }

        //This function is for displaying which event is being called and the counter number
        public static void IncrementAndDisplay(string functionName)
        {
            globalCounter++;
            string conCat = "";
            if (eventCallSequence != "")
            {
                conCat = " -> ";
            }
            eventCallSequence = eventCallSequence + conCat + functionName;
            System.Windows.Forms.MessageBox.Show($"Called by: {functionName} \n GlobalCounter is now: {globalCounter} \n EventCallSequence is: {eventCallSequence}");
        }

    }
}