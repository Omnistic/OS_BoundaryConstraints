using System;
using ZOSAPI;
using ZOSAPI.Editors.LDE;
using ZOSAPI.Editors.MFE;

namespace CSharpUserExtensionApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }
            
            BeginUserExtension();
        }

        static void BeginUserExtension()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Plugin)
            {
                HandleError("User plugin was started in the wrong mode: expected Plugin, found " + TheApplication.Mode.ToString());
                return;
            }
			
            // Chech the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }

            TheApplication.ProgressPercent = 0;
            TheApplication.ProgressMessage = "Running Extension...";

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;
			if (!TheApplication.TerminateRequested) // This will be 'true' if the user clicks on the Cancel button
            {
                // Add your custom code here...

                // The Lens Data Editor
                ILensDataEditor TheLDE = TheSystem.LDE;

                // The Merit Function Editor
                IMeritFunctionEditor TheMFE = TheSystem.MFE;

                // Initialize MF operand
                IMFERow Operand;

                // Loop over the surfaces (except last one)
                int NumberOfSurfaces = TheLDE.NumberOfSurfaces;
                int OperandPos = 1;

                for (int SurfaceID = 0; SurfaceID < NumberOfSurfaces - 1; SurfaceID++)
                {
                    // Retrieve current surface
                    ILDERow CurrentSurface = TheLDE.GetSurfaceAt(SurfaceID);

                    if (double.IsInfinity(CurrentSurface.Thickness) || CurrentSurface.ThicknessCell.GetSolveData().Type != ZOSAPI.Editors.SolveType.Variable)
                    {
                        continue;
                    }

                    // Test if material cell is empty (air)
                    if (CurrentSurface.Material == "")
                    {
                        // Insert MNCA
                        Operand = TheMFE.InsertNewOperandAt(OperandPos);
                        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNCA);
                        Operand.Weight = 1;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = SurfaceID;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = SurfaceID;
                        OperandPos++;

                        // Insert MXCA
                        Operand = TheMFE.InsertNewOperandAt(OperandPos);
                        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MXCA);
                        Operand.Target = 1000;
                        Operand.Weight = 1;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = SurfaceID;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = SurfaceID;
                        OperandPos++;

                        // Insert MNEA
                        Operand = TheMFE.InsertNewOperandAt(OperandPos);
                        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNEA);
                        Operand.Weight = 1;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = SurfaceID;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = SurfaceID;
                        OperandPos++;
                    }
                    else if (CurrentSurface.Material != "MIRROR")
                    {
                        // Insert MNCG
                        Operand = TheMFE.InsertNewOperandAt(OperandPos);
                        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNCG);
                        Operand.Weight = 1;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = SurfaceID;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = SurfaceID;
                        OperandPos++;

                        // Insert MXCG
                        Operand = TheMFE.InsertNewOperandAt(OperandPos);
                        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MXCG);
                        Operand.Target = 1000;
                        Operand.Weight = 1;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = SurfaceID;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = SurfaceID;
                        OperandPos++;

                        // Insert MNEG
                        Operand = TheMFE.InsertNewOperandAt(OperandPos);
                        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNEG);
                        Operand.Weight = 1;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param1).IntegerValue = SurfaceID;
                        Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Param2).IntegerValue = SurfaceID;
                        OperandPos++;
                    }
                }

                // Write comment
                Operand = TheMFE.InsertNewOperandAt(1);
                Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.BLNK);
                Operand.GetOperandCell(ZOSAPI.Editors.MFE.MeritColumn.Comment).Value = "Individual air and glass thickness boundary constraints.";
            }
			
			
			// Clean up
            FinishUserExtension(TheApplication);
        }
		
		static void FinishUserExtension(IZOSAPI_Application TheApplication)
		{
            // Note - OpticStudio will stay in User Extension mode until this application exits
			if (TheApplication != null)
			{
                TheApplication.ProgressMessage = "Complete";
                TheApplication.ProgressPercent = 100;
			}
		}

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
