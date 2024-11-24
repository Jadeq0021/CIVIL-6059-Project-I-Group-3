using Autodesk.Revit.DB;

//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Controls;
using System.Windows;
using Autodesk.Revit.UI;
using System.Xaml.Schema;
using System.Drawing.Drawing2D;

namespace CodeAgent
{
    internal class Utils
    {
        public static Floor CreateFloor(Document document, List<UV> points, Level level, string name)
        {
            // Convert UV points to XYZ points
            List<XYZ> xyzPoints = points.Select(p => new XYZ(p.U, p.V, 0)).ToList();

            // Create curves from the XYZ points
            List<Curve> curves = new List<Curve>();
            for (int i = 0; i < xyzPoints.Count; i++)
            {
                int nextIndex = (i + 1) % xyzPoints.Count;
                Curve line = Line.CreateBound(xyzPoints[i], xyzPoints[nextIndex]);
                curves.Add(line);
            }

            // Create a curve loop from the curves
            CurveArray curveLoop = new CurveArray();
            // 遍历添加曲线
            foreach (Curve curve in curves)
            {
                curveLoop.Append(curve);
            }
            // Get the floor type
            FloorType floorType = new FilteredElementCollector(document)
                .OfClass(typeof(FloorType))
                .WhereElementIsElementType()
                .Cast<FloorType>()
                .FirstOrDefault(x => x.Name == "常规 - 150mm"); // Set your own selected floor type in your Revit

            if (floorType == null)
            {
                throw new InvalidOperationException("找不到地板类型 '常规 - 150mm'");
            }

            Floor floor = null;
            using (Transaction tr = new Transaction(document, "Create Floor"))
            {
                tr.Start();
                floor = document.Create.NewFloor(curveLoop , floorType, level, false);
                if (floor != null)
                {
                    floor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(name);
                }
                tr.Commit();
            }

            return floor;
        }
    

    public static void SetWallThick(Document doc, WallType wType, double wThick)
        {
            Transaction tra = new Transaction(doc, "修改墙体厚度");
            tra.Start();
            CompoundStructure cs = wType.GetCompoundStructure();
            //获取所有层
            IList<CompoundStructureLayer> lstLayers = cs.GetLayers();
            foreach (CompoundStructureLayer item in lstLayers)
            {
                if (item.Function == MaterialFunctionAssignment.Structure)
                {//这里只考虑有一个结构层，如果有多个就自己算算
                    item.Width = wThick;
                    break;
                }
            }
            //修改后要设置一遍
            cs.SetLayers(lstLayers);
            wType.SetCompoundStructure(cs);
            tra.Commit();

        }

        public static double CalculateSideLength(double absoluteArea, double wallThickness)
        {

            double d = absoluteArea + 0.25 * wallThickness * wallThickness;

            // 计算判别式  
            double discriminant = Math.Sqrt(d) + 2.5*wallThickness;


            // 输出正数解（如果有的话）  
            return (discriminant- wallThickness)/2;
            
          
        }

    // determine whether two 3-dimensional points are equal
        public static bool IsEqual(XYZ point_1, XYZ point_2)
        {
            if (point_1.X == point_2.X && point_1.Y == point_2.Y && point_1.Z == point_2.Z)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        // calculate the midpoint of a line consisting of two 2-dimensional points
        public static UV MidPointForLine(UV point_1, UV point_2)
        {
            UV midPoint = new UV((point_1.U + point_2.U) / 2, (point_1.V + point_2.V) / 2);
            return midPoint;
        }

        // calculate the midpoint of a rectangle consisting of four 2-dimensional points
        public static UV MidPointForRectangle(UV point_1, UV point_2, UV point_3, UV point_4)
        {
            UV midPoint = new UV((point_1.U + point_2.U + point_3.U + point_4.U) / 4, (point_1.V + point_2.V + point_3.V + point_4.V) / 4);
            return midPoint;
        }

        // calculate the midpoint of a wall
        public static UV MidPointForWall(Wall wall)
        {
            XYZ startPoint = (wall.Location as LocationCurve).Curve.GetEndPoint(0);
            XYZ endPoint = (wall.Location as LocationCurve).Curve.GetEndPoint(1);
            UV midPoint = new UV((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2);
            return midPoint;
        }

        public static ElementId CreatWallType(Document doc, string wallTypeName, double width)
        {

            ElementId wallTypeId = null;
            FilteredElementCollector Col = new FilteredElementCollector(doc);
            var familySymbolList = Col.OfClass(typeof(WallType)).ToList();
            WallType baseWallType = null;
            WallType newWallType = null;
            using (Transaction transaction = new Transaction(doc))
            {
                if (transaction.Start("创建新墙类型") == TransactionStatus.Started)
                {
                    bool hascreate = false;
                   
                    //获取一个墙类型复制并重命名
                    foreach (WallType item in familySymbolList)
                    {
                        if (item.Category.Name == "墙" && item.FamilyName == "基本墙" && item.Name == "常规 - 200mm")
                        {
                            baseWallType = item;
                            continue;
                        }
                    }
                    foreach (WallType item in familySymbolList)
                    {
                        if (item.Name == wallTypeName)
                        {
                            transaction.Commit();
                            wallTypeId = baseWallType.Id;
                            return wallTypeId;
                        }
                    }
                        newWallType = baseWallType.Duplicate(wallTypeName) as WallType;
                        doc.Regenerate();
                        CompoundStructure wallTypeStructure = newWallType.GetCompoundStructure();
                        int endIndex = wallTypeStructure.GetLastCoreLayerIndex();
                        wallTypeStructure.SetLayerWidth(endIndex, width / 304.8);
                        newWallType.SetCompoundStructure(wallTypeStructure);
                        transaction.Commit();
                       
                }
            }
            wallTypeId = newWallType.Id;
            return wallTypeId;
        }

        public static Wall CreateWall(Document document, UV point_1, UV point_2, Level level, string name)
        {
            WallType wallType = new FilteredElementCollector(document)
                .OfClass(typeof(WallType))
                .WhereElementIsElementType()
                .Cast<WallType>()
                .FirstOrDefault(x => x.Name == "常规 - 200mm"); // Set your own selected wall type in your Revit
            Wall wall = null;
            using (Transaction tr = new Transaction(document, "create_wall"))
            {
                tr.Start();
                wall = Wall.Create(document, Line.CreateBound(new XYZ(point_1.U, point_1.V, 0), new XYZ(point_2.U, point_2.V, 0)), level.Id, false);
                wall.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).Set(1);
                wall.WallType = wallType;
                wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(name);
                tr.Commit();
            }
            return wall;
        }

        // create a wall with two 2-dimensional boundary points of the wall
        public static Wall CreateWall(Document document, UV point_1, UV point_2, Level level, string name, string wallThickness_string)
        {
            //bug Wall-Ret_300Con
            WallType wallType = new FilteredElementCollector(document)
                .OfClass(typeof(WallType))
                .WhereElementIsElementType()
                .Cast<WallType>()
                .FirstOrDefault(x => x.Name == wallThickness_string); // Set your own selected wall type in your Revit
            Wall wall = null;
            
            using (Transaction tr = new Transaction(document, "create_wall"))
            {
                tr.Start();

                wall = Wall.Create(document, Line.CreateBound(new XYZ(point_1.U, point_1.V, 0), new XYZ(point_2.U, point_2.V, 0)), level.Id, false);
                wall.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).Set(1);
                wall.WallType = wallType;
                wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(name);
                //厚度时类型属性，建议新建一个墙体类型
                tr.Commit();
            }
            return wall;
        }

        public static WallType CreateOrGetWallType(Document document, double wallThickness, string typeName)
        {
            WallType newType = null;
            // Check if the wall type already exists
            WallType existingType = new FilteredElementCollector(document)
                .OfClass(typeof(WallType))
                .WhereElementIsElementType()
                .Cast<WallType>()
                .FirstOrDefault(x => x.Name == typeName);

            if (existingType != null)
            {
                return existingType;
            }

            // If not, create a new wall type
            WallType baseType = new FilteredElementCollector(document)
                .OfClass(typeof(WallType))
                .WhereElementIsElementType()
                .Cast<WallType>()
                .FirstOrDefault(x => x.Name == "常规 - 200mm");

            if (baseType == null)
            {
                throw new InvalidOperationException("找不到基础墙类型 '常规 - 200mm'");
            }

            using (Transaction transaction = new Transaction(document, "Create Wall Type"))
            {
                transaction.Start();
                newType = baseType.Duplicate(typeName) as WallType;
                if (newType != null)
                {
                    CompoundStructure structure = newType.GetCompoundStructure();
                    structure.SetLayerWidth(structure.GetLastCoreLayerIndex(), wallThickness/304.8);
                    newType.SetCompoundStructure(structure);
                }
                transaction.Commit();
            }

            return newType;
        }

        public static Wall CreateWall(Document document, UV point1, UV point2, Level level, string name, double wallThickness)
        {
            string typeName = $"自定义墙 - {wallThickness}mm";
            WallType wallType = CreateOrGetWallType(document, wallThickness, typeName);

            Wall wall = null;
            using (Transaction transaction = new Transaction(document, "Create Wall"))
            {
                transaction.Start();
                wall = Wall.Create(document, Line.CreateBound(new XYZ(point1.U, point1.V, 0), new XYZ(point2.U, point2.V, 0)), level.Id, false);
                if (wall != null)
                {
                    wall.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).Set(1);
                    wall.WallType = wallType;
                    wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(name);
                }
                transaction.Commit();
            }

            return wall;
        }
    

    // delete a wall given its 2-dimensional start point and end point
    public static void DeleteWall(Document document, UV startPoint, UV endPoint)
        {
            XYZ startPoint_3d = new XYZ(startPoint.U, startPoint.V, 0);
            XYZ endPoint_3d = new XYZ(endPoint.U, endPoint.V, 0);
            Wall wall = new FilteredElementCollector(document)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .FirstOrDefault(x => IsEqual((x.Location as LocationCurve).Curve.GetEndPoint(0), startPoint_3d) && IsEqual((x.Location as LocationCurve).Curve.GetEndPoint(1), endPoint_3d) || IsEqual((x.Location as LocationCurve).Curve.GetEndPoint(1), startPoint_3d) && IsEqual((x.Location as LocationCurve).Curve.GetEndPoint(0), endPoint_3d));
            if (wall != null)
            {
                using (Transaction tr = new Transaction(document, "delete_wall"))
                {
                    tr.Start();
                    document.Delete(wall.Id);
                    tr.Commit();
                }
            }
        }

        // create a floor with four 2-dimensional boundary points
        public static Floor CreateFloor(Document document, UV point_1, UV point_2, UV point_3, UV point_4, Level level, string name)
        {
            XYZ point_1_3d = new XYZ(point_1.U, point_1.V, 0);
            XYZ point_2_3d = new XYZ(point_2.U, point_2.V, 0);
            XYZ point_3_3d = new XYZ(point_3.U, point_3.V, 0);
            XYZ point_4_3d = new XYZ(point_4.U, point_4.V, 0);
            Curve line_1 = Line.CreateBound(point_1_3d, point_2_3d);
            Curve line_2 = Line.CreateBound(point_2_3d, point_3_3d);
            Curve line_3 = Line.CreateBound(point_3_3d, point_4_3d);
            Curve line_4 = Line.CreateBound(point_4_3d, point_1_3d);
            CurveArray curveArray = new CurveArray();
            curveArray.Append(line_1);
            curveArray.Append(line_2);
            curveArray.Append(line_3);
            curveArray.Append(line_4);
            List<Curve> curves = new List<Curve>() { line_1, line_2, line_3, line_4 };
            CurveLoop curveLoop = CurveLoop.Create(curves);
            IList<CurveLoop> curveLoop_l = new List<CurveLoop>() { curveLoop };
            FloorType floorType = new FilteredElementCollector(document)
                .OfClass(typeof(FloorType))
                .WhereElementIsElementType()
                .Cast<FloorType>()
                .FirstOrDefault(x => x.Name == "常规 - 150mm"); // Set your own selected floor type in your Revit
            Floor floor = null;
            using (Transaction tr = new Transaction(document, "create_floor"))
            {
                tr.Start();
                //bug
                floor = document.Create.NewFloor(curveArray, floorType, level,false);
                floor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(name);
                tr.Commit();
            }
            return floor;
        }

        // create a door with one 2-dimensional point and the wall that the door is on
        public static FamilyInstance CreateDoor(Document document, UV point, Wall wall, Level level)
        {
            XYZ point_3d = new XYZ(point.U, point.V, 0);
            FamilySymbol symbol = new FilteredElementCollector(document)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .FirstOrDefault(x => x.FamilyName == "单扇 - 与墙齐"); // Set your own selected door type in your Revit
            FamilyInstance door = null;
            using (Transaction tr = new Transaction(document, "create_door"))
            {
                tr.Start();
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                }
                door = document.Create.NewFamilyInstance(point_3d, symbol, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                tr.Commit();
            }
            return door;
        }

        // create a window with one 2-dimensional point and the wall that the window is on
        public static FamilyInstance CreateWindow(Document document, UV point, Wall wall, Level level)
        {
            double height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
            double window_height = height / 2;
            XYZ point_3d = new XYZ(point.U, point.V, window_height);
            FamilySymbol symbol = new FilteredElementCollector(document)
                .OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .FirstOrDefault(x => x.FamilyName == "固定"); // Set your own selected window type in your Revit
            FamilyInstance window = null;
            using (Transaction tr = new Transaction(document, "create_window"))
            {
                tr.Start();
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                }
                window = document.Create.NewFamilyInstance(point_3d, symbol, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                tr.Commit();
            }
            return window;
        }

        // merge the areas in two modules into one area in one module
        public static Module MergeModule(Document document, Module module_1, Module module_2, Level level, string wallThickness_string)
        {
            UV point_1 = module_1.GetSouthwestPoint();
            UV point_2 = module_1.GetSoutheastPoint();
            UV point_3 = module_1.GetNortheastPoint();
            UV point_4 = module_1.GetNorthwestPoint();
            UV point_5 = module_2.GetSouthwestPoint();
            UV point_6 = module_2.GetSoutheastPoint();
            UV point_7 = module_2.GetNortheastPoint();
            UV point_8 = module_2.GetNorthwestPoint();
            if (point_1.V == point_7.V)
            {
                CreateMergedWalls(document, point_1, point_2, point_7, point_8, level, wallThickness_string);
            }
            else if (point_3.V == point_5.V)
            {
                CreateMergedWalls(document, point_3, point_4, point_5, point_6, level, wallThickness_string);
            }
            else if (point_2.U == point_8.U)
            {
                CreateMergedWalls(document, point_2, point_3, point_8, point_5, level, wallThickness_string);
            }
            else if (point_4.U == point_6.U)
            {
                CreateMergedWalls(document, point_4, point_1, point_6, point_7, level, wallThickness_string);
            }
            return new Module(document, module_1, module_2);
        }

        // create new walls for the merged area
        public static void CreateMergedWalls(Document document, UV point_1, UV point_2, UV point_3, UV point_4, Level level, string wallThickness_string)
        {
            DeleteWall(document, point_1, point_2);
            DeleteWall(document, point_3, point_4);
            if (point_1.U == point_2.U)
            {
                List<UV> points = new List<UV>() { point_1, point_2, point_3, point_4 };
                points.Sort((x, y) => x.V.CompareTo(y.V));
                if (points[0].V != points[1].V)
                {
                    CreateWall(document, points[0], points[1], level, "merged_wall_1", wallThickness_string);
                }
                if (points[2].V != points[3].V)
                {
                    CreateWall(document, points[2], points[3], level, "merged_wall_2", wallThickness_string);
                }
            }
            else if (point_1.V == point_2.V)
            {
                List<UV> points = new List<UV>() { point_1, point_2, point_3, point_4 };
                points.Sort((x, y) => x.U.CompareTo(y.U));
                if (points[0].U != points[1].U)
                {
                    CreateWall(document, points[0], points[1], level, "merged_wall_1", wallThickness_string);
                }
                if (points[2].U != points[3].U)
                {
                    CreateWall(document, points[2], points[3], level, "merged_wall_2", wallThickness_string);
                }
            }
        }

        // find the shortest shared wall between two rooms
        public static Wall FindSharedWall(Document document, Room room_1, Room room_2, Level level)
        {
            UV point_1 = room_1.GetSouthwestPoint();
            UV point_2 = room_1.GetSoutheastPoint();
            UV point_3 = room_1.GetNortheastPoint();
            UV point_4 = room_1.GetNorthwestPoint();
            UV point_5 = room_2.GetSouthwestPoint();
            UV point_6 = room_2.GetSoutheastPoint();
            UV point_7 = room_2.GetNortheastPoint();
            UV point_8 = room_2.GetNorthwestPoint();

            Wall overlap_wall_1 = WallofOverlappedRooms(document, point_1, point_2, point_3, point_4, point_5, point_6, point_7, point_8);
            Wall overlap_wall_2 = WallofOverlappedRooms(document, point_5, point_6, point_7, point_8, point_1, point_2, point_3, point_4);
            if(overlap_wall_1 != null)
            {
                return overlap_wall_1;
            }
            if(overlap_wall_2 != null)
            {
                return overlap_wall_2;
            }

            if (point_1.V == point_7.V)
            {
                return FindShortestSharedWall(document, point_1, point_2, point_7, point_8, level);
            }
            else if (point_3.V == point_5.V)
            {
                return FindShortestSharedWall(document, point_3, point_4, point_5, point_6, level);
            }
            else if (point_2.U == point_8.U)
            {
                return FindShortestSharedWall(document, point_2, point_3, point_8, point_5, level);
            }
            else if (point_4.U == point_6.U)
            {
                return FindShortestSharedWall(document, point_4, point_1, point_6, point_7, level);
            }
            else
            {
                return null;
            }
        }

        // determine whether the areas of two rooms overlap, and if so, return the overlapping wall
        public static Wall WallofOverlappedRooms(Document document, UV point_1, UV point_2, UV point_3, UV point_4, UV point_5, UV point_6, UV point_7, UV point_8)
        {
            if(point_1.U > point_5.U || point_2.U < point_6.U || point_1.V > point_5.V || point_4.V < point_8.V)
            {
                return null;
            }
            else
            {
                if(point_1.U < point_5.U && point_1.V < point_5.V)
                {
                    return GetWallByTwoPoints(document, point_5, point_6);
                }
                if(point_2.U > point_6.U && point_2.V < point_6.V)
                {
                    return GetWallByTwoPoints(document, point_5, point_6);
                }
                if(point_4.U < point_8.U && point_4.V > point_8.V)
                {
                    return GetWallByTwoPoints(document, point_7, point_8);
                }
                if(point_3.U > point_7.U && point_3.V > point_7.V)
                {
                    return GetWallByTwoPoints(document, point_7, point_8);
                }
                return null;
            }
        }

        // find the shortest shared wall between two modules given four points
        public static Wall FindShortestSharedWall(Document document, UV point_1, UV point_2, UV point_3, UV point_4, Level level)
        {
            if (point_1.U == point_2.U)
            {
                List<UV> points = new List<UV>() { point_1, point_2, point_3, point_4 };
                points.Sort((x, y) => x.V.CompareTo(y.V));
                return GetWallByTwoPoints(document, points[1], points[2]);
            }
            else if (point_1.V == point_2.V)
            {
                List<UV> points = new List<UV>() { point_1, point_2, point_3, point_4 };
                points.Sort((x, y) => x.U.CompareTo(y.U));
                return GetWallByTwoPoints(document, points[1], points[2]);
            }
            else
            {
                return null;
            }
        }

        // get the wall based on two points
        public static Wall GetWallByTwoPoints(Document document, UV point_1, UV point_2)
        {
            Curve curve = Line.CreateBound(new XYZ(point_1.U, point_1.V, 0), new XYZ(point_2.U, point_2.V, 0));
            List<Wall> walls = new FilteredElementCollector(document)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();
            foreach (Wall wall in walls)
            {
                Curve wall_curve = (wall.Location as LocationCurve).Curve;
                if (curve.Intersect(wall_curve).ToString() == "Equal")
                {
                    return wall;
                }
            }
            return null;
        }

        // create a door between two rooms
        public static FamilyInstance CreateDoorBetweenRooms(Document document, Room room_1, Room room_2, Level level)
        {
            UV point_1 = room_1.GetSouthwestPoint();
            UV point_2 = room_1.GetSoutheastPoint();
            UV point_3 = room_1.GetNortheastPoint();
            UV point_4 = room_1.GetNorthwestPoint();
            UV point_5 = room_2.GetSouthwestPoint();
            UV point_6 = room_2.GetSoutheastPoint();
            UV point_7 = room_2.GetNortheastPoint();
            UV point_8 = room_2.GetNorthwestPoint();
            UV midpoint;
            if (point_1.V == point_7.V)
            {
                if(point_2.U - point_1.U > point_7.U - point_8.U)
                {
                    midpoint = MidPointForLine(point_7, point_8);
                }
                else
                {
                    midpoint = MidPointForLine(point_1, point_2);
                }
                return CreateDoor(document, midpoint, FindSharedWall(document, room_1, room_2, level), level);
            }
            else if (point_3.V == point_5.V)
            {
                if (point_3.U - point_4.U > point_6.U - point_5.U)
                {
                    midpoint = MidPointForLine(point_5, point_6);
                }
                else
                {
                    midpoint = MidPointForLine(point_3, point_4);
                }
                return CreateDoor(document, midpoint, FindSharedWall(document, room_1, room_2, level), level);
            }
            else if (point_2.U == point_8.U)
            {
                if(point_3.V - point_2.V > point_8.V - point_5.V)
                {
                    midpoint = MidPointForLine(point_5, point_8);
                }
                else
                {
                    midpoint = MidPointForLine(point_2, point_3);
                }
                return CreateDoor(document, midpoint, FindSharedWall(document, room_1, room_2, level), level);
            }
            else if (point_4.U == point_6.U)
            {
                if(point_4.V - point_1.V > point_7.V - point_6.V)
                {
                    midpoint = MidPointForLine(point_6, point_7);
                }
                else
                {
                    midpoint = MidPointForLine(point_1, point_4);
                }
                return CreateDoor(document, midpoint, FindSharedWall(document, room_1, room_2, level), level);
            }
            else
            {
                return null;
            }
        }

        // create a door or window between a room and the external space
        public static FamilyInstance CreateDoorOrWindowBetweenRoomAndExternalSpace(Document document, Room room, Level level, string direction, string type)
        {
            Wall wall = null;
            if (direction == "north")
            {
                wall = room.GetNorthWall();
            }
            else if (direction == "south")
            {
                wall = room.GetSouthWall();
            }
            else if (direction == "east")
            {
                wall = room.GetEastWall();
            }
            else if (direction == "west")
            {
                wall = room.GetWestWall();
            }
            if (type == "door")
            {
                return CreateDoor(document, MidPointForWall(wall), wall, level);
            }
            else if (type == "window")
            {
                return CreateWindow(document, MidPointForWall(wall), wall, level);
            }
            else
            {
                return null;
            }
        }
    }
}
