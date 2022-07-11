using System;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Timeliner;

namespace Timeliner_Part1
{
    [PluginAttribute("CreateTask", "TwentyTwo",
        DisplayName = "Create TimelinerTask", ToolTip = "Creating timeliner tasks from selection sets.")]

    public class MainClass : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            // current document
            Document doc = Application.ActiveDocument;
            // get timeliner obj
            DocumentTimeliner timeliner = doc.GetTimeliner();
            // collect selection sets
            SavedItemCollection selectionSets = doc.SelectionSets.RootItem.Children;
            // int value
            int index = 1;
            // for each selection set
            foreach (SavedItem set in selectionSets)
            {
                // create timeliner task with display name
                // startdate, enddate & tasktype properties
                TimelinerTask task = new TimelinerTask()
                {
                    DisplayName = set.DisplayName,
                    PlannedStartDate = new DateTime(2022, index , 7),
                    PlannedEndDate = new DateTime(2022, index + 1, 7),
                    SimulationTaskTypeName = "Construct"
                };

                // saveditem to selection set  
                SelectionSet selSect = set as SelectionSet;
                // create selection source from selection set
                SelectionSource selSource = doc.SelectionSets.CreateSelectionSource(selSect);
                // create selection source collection from selection source
                SelectionSourceCollection selSourceCol = new SelectionSourceCollection() { selSource };
                // attach selection soruce colleciton to timeliner task
                task.Selection.CopyFrom(selSourceCol);
                // add timeliner task to document's timeliner obj
                timeliner.TaskAddCopy(task);
                // increment of 1
                index++;
            } 
            return 0;
        }
    }
}
