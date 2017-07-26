namespace CanBIM
{
	using System;
	using System.Linq;

	using Autodesk.Revit.DB;
	using Autodesk.Revit.UI;
	
	public partial class ThisApplication
	{
		/// <summary>
		/// Renames all sheets that contain exactly one floor plan to match the name of that view
		/// Copy the whole method into the macro file on your machine,
		/// or download this file and add it to the solution in the macro editor
		/// </summary>
		public void RenameSheetsToMatchViewNames()
		{
			UIDocument uidoc = this.ActiveUIDocument;
			Document doc = uidoc.Document;
			
			// Retrieve all sheets, cast them to type ViewSheet and save them in a list
			// It is important to use ToList or ToElements when modifying elements
			// from a FilteredElementCollector
			var sheets = new FilteredElementCollector(doc)
				.OfClass(typeof(ViewSheet))
				.Cast<ViewSheet>().ToList();
			
			// Initialize a counter that will identify how many sheets have been modified
			int renamedSheetsCount = 0;
			
			// Open a transaction that allows changes to be made to the model
			using (Transaction trans = new Transaction(doc, "Rename sheets"))
			{
				// Start the transaction
				trans.Start();
				
				// Loop through each sheet retrieved above
				foreach (var sheet in sheets)
				{
					// Get all the views placed on the sheet
					var viewsOnSheet = sheet.GetAllPlacedViews();
					
					// If there is not exactly 1 view on the sheet,
					// Skip this sheet and continue to the next one
					if (viewsOnSheet.Count != 1)
					{
						continue;
					}
					
					// Get the id of the view
					// The single call is used because we know there is only 1
					// based on the check above
					var viewId = viewsOnSheet.Single();
					
					// Get the view from the id, and cast it to type View
					var view = (View)doc.GetElement(viewId);
					
					// If the view is not a floor plan,
					// skip the sheet and continue to the next one
					if (view.ViewType != ViewType.FloorPlan)
					{
						continue;
					}
					
					// Open a sub-transaction for the upcoming modification to the sheet
					using (SubTransaction subTrans = new SubTransaction(doc))
					{
						// Start the sub-transaction
						subTrans.Start();
						
						// Set the sheet name to the view name
						sheet.Name = view.Name;
						
						// Commit the change made
						subTrans.Commit();
						
						// Increment the sheet counter by 1
						renamedSheetsCount++;
					}
					
					// End of foreach loop
				}
				
				// Commit the transaction
				trans.Commit();
			}
			
			// Show a message indicating how many sheets were modified
			TaskDialog.Show("Done", "Renamed " + renamedSheetsCount + " sheets");
		}
	}
}
