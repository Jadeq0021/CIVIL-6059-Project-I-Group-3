using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;


namespace CodeAgent
{
    [RegenerationAttribute(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]

    public class MainClassEvent : IExternalEventHandler
    {
        //-----------------------------参数表---------------------------------

        public double model = 0;
        public double area = 40 * 3.2808399 * 3.2808399;
        // UnitUtils.Convert (11.7,UnitTypeId.Feet,UnitTypeId.Millimeters);
       // public double wallThickness = UnitUtils.Convert(0.2, DisplayUnitType.DUT_METERS, DisplayUnitType.DUT_DECIMAL_FEET);
        public double wallThickness = 0.2;
        public double wide = 0.2;
        public double height = 0.2;

        //-------------------------------------------------------------
        Autodesk.Revit.ApplicationServices.Application application = null; //应用程序
        UIDocument uIDocument = null; //ui文档
        Autodesk.Revit.DB.Document doc = null;

        public void Execute(UIApplication uIApp)
        {
            Autodesk.Revit.ApplicationServices.Application application = uIApp.Application;
            UIDocument uIDoc = uIApp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uIDoc.Document;
            View view = uIDoc.ActiveView;
            Level level = new FilteredElementCollector(doc)
               .OfClass(typeof(Level))
               .WhereElementIsNotElementType()
               .Cast<Level>()
               .FirstOrDefault(x => x.Name == "标高 1");
            if (model == 0)
            {
                area = area * 3.2808399 * 3.2808399;
                string wallThickness_string = (wallThickness * 1000).ToString();
                ElementId wallTypeId = Utils.CreatWallType(doc, "常规 - " + wallThickness_string.ToString() + "mm", wallThickness * 1000);
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);
                //TaskDialog.Show("Selection", "x,y,z=" + point.ToString());
                wallThickness *= 3.2808399;
                wallThickness_string = "常规 - " + wallThickness_string.ToString() + "mm";

           

                double length = Utils.CalculateSideLength(area, wallThickness);
                //TaskDialog.Show("Hello World", length.ToString());
                // example_1
                // 2-dimensional boundary points of module_1
                // Define points for the first module
                UV point_1 = pushpoint + new UV(0, 0);
                UV point_2 = pushpoint + new UV(length, 0);
                UV point_3 = pushpoint + new UV(length, length * 2);
                UV point_4 = pushpoint + new UV(0, length * 2);

                // Define points for the second module
                UV point_5 = pushpoint + new UV(length, 0);
                UV point_6 = pushpoint + new UV(length * 2, 0);
                UV point_7 = pushpoint + new UV(length * 2, length * 2);
                UV point_8 = pushpoint + new UV(length, length * 2);
                Module module_1 = new Module(doc, point_1, point_2, point_3, point_4, level, "module_1", wallThickness_string);
                Module module_2 = new Module(doc, point_5, point_6, point_7, point_8, level, "module_2", wallThickness_string);
                // create rooms in the modules
                Room bedroom = new Room(doc, module_1, level, "Bedroom");
                Room living_room = new Room(doc, module_2, level, "Living room");
                // create a door between the bedroom and the living room
                FamilyInstance Door_1 = Utils.CreateDoorBetweenRooms(doc, bedroom, living_room, level);
                // create a door between the living room and the external space
                FamilyInstance Door_2 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, living_room, level, "south", "door");
                // create a window between the living room and the external space
                FamilyInstance Window_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, living_room, level, "north", "window");
                // create a window between the bedroom and the external space
                FamilyInstance Window_2 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, bedroom, level, "north", "window");


            }
            else if(model == 10)
            {
                // Create
                UV MODULE_A1_Insert = new UV(0, 0);
                List<UV> MODULE_A1 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((2450+125/2+85/2)/304.8, -(0)/304.8),
            new UV((2450 + 125 / 2 + 85 / 2)/304.8, -(3610 + 85)/304.8),
            new UV((696.01)/304.8, -(3610+85)/304.8),
            new UV((696.01)/304.8, -(2625.38+85/2)/304.8),
            new UV((241-85/2-35)/304.8, -(2625.38+85/2)/304.8),
            new UV((241-85/2-35)/304.8, -(1093.73+85)/304.8),
            new UV((0)/304.8, -(1093.73+85)/304.8)
            };
                List<double> MODULE_A1_wall = new List<double> { 85, 125, 85, 85, 85, 85, 85, 125 };
                Module module_1 = new Module(doc, MODULE_A1, level, "MODULE A1", MODULE_A1_wall, MODULE_A1_Insert);
                Room room_module_1 = new Room(doc, module_1, level, "MODULE A1");
                FamilyInstance Window_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_1, level, "west", "window");


                // Create modules
                UV MODULE_A2_Insert = new UV((-2797.7947 + 62.5) / 304.8, (-1516.8384 - 62.5) / 304.8);
                List<UV> MODULE_A2 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((2710+125/2+85/2)/304.8, -(0)/304.8),
            new UV((2710+125/2+85/2)/304.8, -(1023 + 85)/304.8),
            new UV((2339.78+125/2+85/2)/304.8, -(1023 + 85)/304.8),
            new UV((2339.78+125/2+85/2)/304.8, -(2050 + 85)/304.8),
            new UV((0)/304.8, -(2050 + 85)/304.8),
            };
                List<double> MODULE_A2_wall = new List<double> { 85, 85, 85, 85, 85, 125 };
                Module module_2 = new Module(doc, MODULE_A2, level, "MODULE A2", MODULE_A2_wall, MODULE_A2_Insert);
                Room room_module_2 = new Room(doc, module_2, level, "MODULE A2");
                FamilyInstance Window_2 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_2, level, "west", "window");


                UV MODULE_A3_Insert = new UV((652.6425 + 42.5) / 304.8, (-3750.8431 - 125) / 304.8);
                List<UV> MODULE_A3 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((1775+85/2+125/2)/304.8, -(0)/304.8),
            new UV((1775+85/2+125/2)/304.8, -(2050 + 250/2+250/2)/304.8),
            new UV((0)/304.8, -(2050 + 250/2+250/2)/304.8),
            };
                List<double> MODULE_A3_wall = new List<double> { 250, 125, 250, 85 };
                Module module_3 = new Module(doc, MODULE_A3, level, "MODULE A3", MODULE_A3_wall, MODULE_A3_Insert);
                Room room_module_3 = new Room(doc, module_3, level, "MODULE A3");
                FamilyInstance Window_3 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_3, level, "east", "window");
                FamilyInstance door_3 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_3, level, "west", "door");



                UV MODULE_A4_Insert = new UV((-2797.7947 + 62.5) / 304.8, (-3750.8431 - 280 / 2) / 304.8);
                List<UV> MODULE_A4 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((2340+125/2+85/2)/304.8, -(0)/304.8),
            new UV((2340+125/2+85/2)/304.8, -(2050 + 275/2+225/2)/304.8),
            new UV((0)/304.8, -(2050 + 275/2+225/2)/304.8),
            };
                List<double> MODULE_A4_wall = new List<double> { 275, 85, 225, 125 };
                Module module_4 = new Module(doc, MODULE_A4, level, "MODULE A4", MODULE_A4_wall, MODULE_A4_Insert);
                Room room_module_4 = new Room(doc, module_4, level, "MODULE A4");
                FamilyInstance Window_4 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_4, level, "west", "window");
                FamilyInstance door_4 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_4, level, "east", "door");



                UV MODULE_A5_Insert = new UV((-2797.794 + 62.5) / 304.8, (-6314.8712 - 50) / 304.8);
                List<UV> MODULE_A5 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((5185+125/2+125/2)/304.8, -(0)/304.8),
            new UV((5185+125/2+125/2)/304.8, -(2830 + 85)/304.8),
            new UV((0)/304.8, -(2830 + 85)/304.8),
            };
                List<double> MODULE_A5_wall = new List<double> { 85, 125, 85, 125 };
                Module module_5 = new Module(doc, MODULE_A5, level, "MODULE A5", MODULE_A5_wall, MODULE_A5_Insert);
                Room room_module_5 = new Room(doc, module_5, level, "MODULE A5");
                FamilyInstance Window_5 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "west", "window");
                FamilyInstance Window_5_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "east", "window");
                FamilyInstance door_5 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "north", "door");
                FamilyInstance door_5_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "south", "door");


                UV MODULE_C_Insert = new UV((-2797.7947 + 175) / 304.8, (-9327.3926 - 175) / 304.8);
                List<UV> MODULE_C = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((1865+350/2)/304.8, -(0)/304.8),
            new UV((1865+350/2)/304.8, -(965 + 350/2+50)/304.8),
            new UV((1453.4+350/2+50)/304.8, -(965 + 350/2+50)/304.8),
            new UV((1453.4+350/2+50)/304.8, -(1865 + 350/2+50)/304.8),
            new UV((0)/304.8, -(1865 + 350/2+50)/304.8),
            };
                List<double> MODULE_C_wall = new List<double> { 350, 100, 100, 85, 100, 350 };
                Module module_C = new Module(doc, MODULE_C, level, "MODULE A5", MODULE_C_wall, MODULE_C_Insert);
                Room room_module_C = new Room(doc, module_C, level, "MODULE A5");



                UV MODULE_A6_Insert = new UV((807.0765 + 75) / 304.8, (-9327.3926 - 75) / 304.8);
                List<UV> MODULE_A6 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((2663+700+50+85/2)/304.8, -(0)/304.8),
            new UV((2663+700+50+85/2)/304.8, -(2065 + 150/2+50)/304.8),
            new UV((0)/304.8, -(2065 + 150/2+50)/304.8),
            };
                List<double> MODULE_A6_wall = new List<double> { 150, 85, 100, 100 };
                Module module_A6 = new Module(doc, MODULE_A6, level, "MODULE A5", MODULE_A6_wall, MODULE_A6_Insert);
                Room room_module_A6 = new Room(doc, module_A6, level, "MODULE A5");
                FamilyInstance door_A6 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_A6, level, "west", "door");


                List<UV> floor_1 = new List<UV> {
            new UV((-248.6)/304.8, -(2710)/304.8),
            new UV((652.6)/304.8, -(2710)/304.8),
            new UV((652.6)/304.8, -(6357)/304.8),
            new UV((-248.6)/304.8, -(6357)/304.8),
            };
                Utils.CreateFloor(doc, floor_1, level, "MODULE A5");

            }
            else if (model == 1)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);

                // Create
                UV MODULE_A1_Insert = pushpoint;
                List<UV> MODULE_A1 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+125/2+85/2)/304.8, -(0)/304.8),
            new UV((wide + 125 / 2 + 85 / 2)/304.8, -(height + 85)/304.8),
            new UV((696.01)/304.8, -(height+85)/304.8),
            new UV((696.01)/304.8, -(2625.38+85/2)/304.8),
            new UV((241-85/2-35)/304.8, -(2625.38+85/2)/304.8),
            new UV((241-85/2-35)/304.8, -(1093.73+85)/304.8),
            new UV((0)/304.8, -(1093.73+85)/304.8)
            };
                List<double> MODULE_A1_wall = new List<double> { 85, 125, 85, 85, 85, 85, 85, 125 };
                Module module_1 = new Module(doc, MODULE_A1, level, "MODULE A1", MODULE_A1_wall, MODULE_A1_Insert);
                Room room_module_1 = new Room(doc, module_1, level, "MODULE A1");
                FamilyInstance Window_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_1, level, "west", "window");


           

            }

            else if (model == 2)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);

           


                // Create modules
                UV MODULE_A2_Insert = pushpoint;
                List<UV> MODULE_A2 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+125/2+85/2)/304.8, -(0)/304.8),
            new UV((wide+125/2+85/2)/304.8, -(1023 + 85)/304.8),
            new UV((2339.78+125/2+85/2)/304.8, -(1023 + 85)/304.8),
            new UV((2339.78+125/2+85/2)/304.8, -(height + 85)/304.8),
            new UV((0)/304.8, -(height + 85)/304.8),
            };
                List<double> MODULE_A2_wall = new List<double> { 85, 85, 85, 85, 85, 125 };
                Module module_2 = new Module(doc, MODULE_A2, level, "MODULE A2", MODULE_A2_wall, MODULE_A2_Insert);
                Room room_module_2 = new Room(doc, module_2, level, "MODULE A2");
                FamilyInstance Window_2 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_2, level, "west", "window");

            }

            else if (model == 3)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);




                UV MODULE_A3_Insert = pushpoint;
                List<UV> MODULE_A3 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+85/2+125/2)/304.8, -(0)/304.8),
            new UV((wide+85/2+125/2)/304.8, -(height + 250/2+250/2)/304.8),
            new UV((0)/304.8, -(height + 250/2+250/2)/304.8),
            };
                List<double> MODULE_A3_wall = new List<double> { 250, 125, 250, 85 };
                Module module_3 = new Module(doc, MODULE_A3, level, "MODULE A3", MODULE_A3_wall, MODULE_A3_Insert);
                Room room_module_3 = new Room(doc, module_3, level, "MODULE A3");
                FamilyInstance Window_3 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_3, level, "east", "window");
                FamilyInstance door_3 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_3, level, "west", "door");




            }

            else if (model == 4)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);





                UV MODULE_A4_Insert = pushpoint;
                List<UV> MODULE_A4 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+125/2+85/2)/304.8, -(0)/304.8),
            new UV((wide+125/2+85/2)/304.8, -(height + 275/2+225/2)/304.8),
            new UV((0)/304.8, -(height + 275/2+225/2)/304.8),
            };
                List<double> MODULE_A4_wall = new List<double> { 275, 85, 225, 125 };
                Module module_4 = new Module(doc, MODULE_A4, level, "MODULE A4", MODULE_A4_wall, MODULE_A4_Insert);
                Room room_module_4 = new Room(doc, module_4, level, "MODULE A4");
                FamilyInstance Window_4 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_4, level, "west", "window");
                FamilyInstance door_4 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_4, level, "east", "door");



            }
            else if (model == 5)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);

            



                UV MODULE_A5_Insert = pushpoint;
                List<UV> MODULE_A5 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+125/2+125/2)/304.8, -(0)/304.8),
            new UV((wide+125/2+125/2)/304.8, -(height + 85)/304.8),
            new UV((0)/304.8, -(height + 85)/304.8),
            };
                List<double> MODULE_A5_wall = new List<double> { 85, 125, 85, 125 };
                Module module_5 = new Module(doc, MODULE_A5, level, "MODULE A5", MODULE_A5_wall, MODULE_A5_Insert);
                Room room_module_5 = new Room(doc, module_5, level, "MODULE A5");
                FamilyInstance Window_5 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "west", "window");
                FamilyInstance Window_5_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "east", "window");
                FamilyInstance door_5 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "north", "door");
                FamilyInstance door_5_1 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_5, level, "south", "door");



            }
            else if (model == 6)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);

                UV MODULE_A6_Insert = pushpoint;
                List<UV> MODULE_A6 = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+700+50+85/2)/304.8, -(0)/304.8),
            new UV((wide+700+50+85/2)/304.8, -(height + 150/2+50)/304.8),
            new UV((0)/304.8, -(height + 150/2+50)/304.8),
            };
                List<double> MODULE_A6_wall = new List<double> { 150, 85, 100, 100 };
                Module module_A6 = new Module(doc, MODULE_A6, level, "MODULE A5", MODULE_A6_wall, MODULE_A6_Insert);
                Room room_module_A6 = new Room(doc, module_A6, level, "MODULE A5");
                FamilyInstance door_A6 = Utils.CreateDoorOrWindowBetweenRoomAndExternalSpace(doc, room_module_A6, level, "west", "door");


          

            }
            else if (model == 7)
            {
                ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
                XYZ point = uIDoc.Selection.PickPoint(snapTypes, "请选择端点或交点");
                UV pushpoint = new UV(point.X, point.Y);

                UV MODULE_C_Insert = pushpoint;
                List<UV> MODULE_C = new List<UV> {
            new UV((0)/304.8, (0)/304.8),
            new UV((wide+350/2)/304.8, -(0)/304.8),
            new UV((wide+350/2)/304.8, -(height-900 + 350/2+50)/304.8),
            new UV((1453.4+350/2+50)/304.8, -(height-900 + 350/2+50)/304.8),
            new UV((1453.4+350/2+50)/304.8, -(height + 350/2+50)/304.8),
            new UV((0)/304.8, -(height + 350/2+50)/304.8),
            };
                List<double> MODULE_C_wall = new List<double> { 350, 100, 100, 85, 100, 350 };
                Module module_C = new Module(doc, MODULE_C, level, "MODULE A5", MODULE_C_wall, MODULE_C_Insert);
                Room room_module_C = new Room(doc, module_C, level, "MODULE A5");



   

            }

        }

        public string GetName()
        {
            return "OpenWindowBeam";
            throw new System.NotImplementedException();
        }

    }
}