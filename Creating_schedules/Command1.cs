#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Controls;

#endregion

namespace Creating_schedules
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            Transaction t = new Transaction(doc);
            t.Start("Schedules");
            FilteredElementCollector rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms);
            List<string> Department = new List<string>();
            foreach (SpatialElement curroom  in rooms)
            {
                string curdep = Getparametervaluebyname(curroom, "Department");
                Department.Add(curdep);

            }

            List<string> Departments = Department.Distinct().ToList();
            int i;
            ElementId catgid = new ElementId(BuiltInCategory.OST_Rooms);
            for (i = 0; i < Departments.Count; i++)
            {
                foreach (SpatialElement curroom in rooms) 
                { 
                    //for (i = 0; i < Departments.Count; i++)
                    //{
                        string curdep = Getparametervaluebyname(curroom, "Department");
                        if (Departments[i] == curdep)
                        {
                            ViewSchedule roomschedule = ViewSchedule.CreateSchedule(doc, catgid);
                            Parameter num = curroom.get_Parameter(BuiltInParameter.ROOM_NUMBER);
                            Parameter nam = curroom.get_Parameter(BuiltInParameter.ROOM_NAME);
                            Parameter dep = curroom.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                            Parameter com = curroom.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                            Parameter lev = curroom.LookupParameter("Level");
                            Parameter area = curroom.get_Parameter(BuiltInParameter.ROOM_AREA);
                            ScheduleField schnum = roomschedule.Definition.AddField(ScheduleFieldType.Instance, num.Id);
                            ScheduleField schnam = roomschedule.Definition.AddField(ScheduleFieldType.Instance, nam.Id);
                            ScheduleField schdep = roomschedule.Definition.AddField(ScheduleFieldType.Instance, dep.Id);
                            ScheduleField schcom = roomschedule.Definition.AddField(ScheduleFieldType.Instance, com.Id);
                            ScheduleField schare = roomschedule.Definition.AddField(ScheduleFieldType.ViewBased, area.Id);
                            ScheduleField schlev = roomschedule.Definition.AddField(ScheduleFieldType.Instance, lev.Id);
                            roomschedule.Name = "Dept-" + curdep.ToString();
                            schlev.IsHidden = true;
                            ScheduleFilter scheduleFilter = new ScheduleFilter(schdep.FieldId, ScheduleFilterType.Equal, Departments[i]);
                            roomschedule.Definition.AddFilter(scheduleFilter);
                            ScheduleSortGroupField sortgr = new ScheduleSortGroupField(schnam.FieldId);
                            ScheduleSortGroupField grpsrt = new ScheduleSortGroupField(schlev.FieldId);
                            grpsrt.ShowHeader = true;
                            grpsrt.ShowFooter = true;
                            grpsrt.ShowBlankLine = true;
                            schare.Definition.ShowGrandTotal = true;
                            schare.DisplayType = ScheduleFieldDisplayType.Totals;
                            roomschedule.Definition.AddSortGroupField(grpsrt);
                            roomschedule.Definition.AddSortGroupField(sortgr);
                            roomschedule.Definition.ShowGrandTotalCount = true;
                            roomschedule.Definition.ShowGrandTotalTitle = true;

                            break;
                            
                        }
                }
            }
            foreach (SpatialElement curroom in rooms)
            {
                ViewSchedule roomscheduledep = ViewSchedule.CreateSchedule(doc, catgid);
                Parameter dep = curroom.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                Parameter area = curroom.get_Parameter(BuiltInParameter.ROOM_AREA);
                ScheduleField schdep = roomscheduledep.Definition.AddField(ScheduleFieldType.Instance, dep.Id);
                ScheduleField schare = roomscheduledep.Definition.AddField(ScheduleFieldType.ViewBased, area.Id);
                roomscheduledep.Name = "Deptartment";
                ScheduleSortGroupField grpsrtdep = new ScheduleSortGroupField(schdep.FieldId);
                roomscheduledep.Definition.AddSortGroupField(grpsrtdep);
                roomscheduledep.Definition.IsItemized = false;
                roomscheduledep.Definition.ShowGrandTotal = true;
                schare.DisplayType = ScheduleFieldDisplayType.Totals;
                roomscheduledep.Definition.ShowGrandTotalTitle = true;
                break;
            }


            t.Commit();
            t.Dispose();


            return Result.Succeeded;
        }

        private string Getparametervaluebyname(Element e, string v)
        {
            IList<Parameter> parameters = e.GetParameters(v);
            Parameter parameter = parameters.First();
            return parameter.AsString();
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
