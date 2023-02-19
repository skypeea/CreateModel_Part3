using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CreateModelPlugin_Part3
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreateModelPlugin_Part3 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);

            List<Level> levels = GetLevelList(doc);
            Level level1 = GetLevelByName(levels, "Уровень 1");
            Level level2 = GetLevelByName(levels, "Уровень 2");
            List<XYZ> points = GetPointsByWidthAndDepth(width, depth);

            List<Wall> wallList = CreateWall(doc, points, level1, level2);

            AddDoor(doc, level1, wallList[0]);
            AddWindow(doc, level1, wallList[1]);
            AddWindow(doc, level1, wallList[2]);
            AddWindow(doc, level1, wallList[3]);

            //AddFootprintRoof(doc, level2, wallList);
            AddExtrusionRoof(doc, level2, width, wallList[1], "Базовая крыша", "Типовой - 400мм", 1500);
            return Result.Succeeded;
        }

        private void AddExtrusionRoof(Document doc, Level level, double length, Wall wall, string typeName, string typeSize, double roofHeight)
        {
            using (Transaction ts = new Transaction(doc, "Create ExtrusionRoof"))
            {
                ts.Start();
                RoofType roofType = new FilteredElementCollector(doc)
                    .OfClass(typeof(RoofType))
                    .OfType<RoofType>()
                    .Where(x => x.FamilyName.Equals(typeName))
                    .Where(x => x.Name.Equals(typeSize))
                    .FirstOrDefault();

                double wallWidth = wall.Width;
                double dt = wallWidth / 2;
                

                double elevation = level.Elevation;

                LocationCurve wallCurve = wall.Location as LocationCurve;
                XYZ point1 = wallCurve.Curve.GetEndPoint(0);
                XYZ pointRoof1 = new XYZ(point1.X + dt, point1.Y - dt, point1.Z + elevation);

                XYZ point2 = wallCurve.Curve.GetEndPoint(1);
                XYZ pointRoof2 = new XYZ(point2.X + dt, point2.Y + dt, point2.Z + elevation);

                XYZ pointMiddle = (point1 + point2) / 2;
                double roofНeightInch = UnitUtils.ConvertToInternalUnits(roofHeight, UnitTypeId.Millimeters);
                XYZ pointMiddleRoof = new XYZ(pointMiddle.X + dt, pointMiddle.Y, pointMiddle.Z + elevation + roofНeightInch);

                CurveArray curveArray = new CurveArray();
                curveArray.Append(Line.CreateBound(pointRoof1, pointMiddleRoof));
                curveArray.Append(Line.CreateBound(pointMiddleRoof, pointRoof2));

                ReferencePlane plane = doc.Create.NewReferencePlane2(pointRoof1, pointRoof2, pointMiddleRoof, doc.ActiveView);
                doc.Create.NewExtrusionRoof(curveArray, plane, level, roofType, 0, -(length+2*dt));
                ts.Commit();
            }
        }

        private void AddFootprintRoof(Document doc, Level level, List<Wall> wallList)
        {
            RoofType roofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Типовой - 400мм"))
                .Where(x => x.FamilyName.Equals("Базовая крыша"))
                .FirstOrDefault();

            double wallWidth = wallList[0].Width;
            double dt = wallWidth / 2;
            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dt, -dt, 0));
            points.Add(new XYZ(dt, -dt, 0));
            points.Add(new XYZ(dt, dt, 0));
            points.Add(new XYZ(-dt, dt, 0));
            points.Add(new XYZ(-dt, -dt, 0));


            Application application = doc.Application;
            CurveArray footprint = application.Create.NewCurveArray();
            for (int i = 0; i < 4; i++)
            {
                LocationCurve curve = wallList[i].Location as LocationCurve;
               
                XYZ p1 = curve.Curve.GetEndPoint(0);
                XYZ p2 = curve.Curve.GetEndPoint(1);
                Line line = Line.CreateBound(p1 + points[i], p2 + points[i + 1]);
                footprint.Append(line);
            }
            using (Transaction ts = new Transaction(doc, "CreateRoof"))
            {
                ts.Start();

                ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();
                FootPrintRoof footPrintRoof = doc.Create.NewFootPrintRoof(footprint, level, roofType, out footPrintToModelCurveMapping);
                //ModelCurveArrayIterator iterator = footPrintToModelCurveMapping.ForwardIterator();
                //iterator.Reset();
                //while(iterator.MoveNext())
                //{
                //    ModelCurve modelCurve = iterator.Current as ModelCurve;
                //    footPrintRoof.set_DefinesSlope(modelCurve, true);
                //    footPrintRoof.set_SlopeAngle(modelCurve, 0.5);
                //}
                foreach (ModelCurve m in footPrintToModelCurveMapping)
                {
                    footPrintRoof.set_DefinesSlope(m, true);
                    footPrintRoof.set_SlopeAngle(m, 0.5);
                }
                ts.Commit();
            }
        }

        private void AddWindow(Document doc, Level level, Wall wall)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            if (windowType != null)
            {
                using (var ts = new Transaction(doc, "create doors"))
                {
                    ts.Start();

                    LocationCurve hostCurve = wall.Location as LocationCurve;
                    XYZ midPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) / 2;

                    if (!windowType.IsActive)
                    {
                        windowType.Activate();
                    }
                    FamilyInstance window = doc.Create.NewFamilyInstance(midPoint, windowType, wall, level, StructuralType.NonStructural);

                    double windowLevelOffset = UnitUtils.ConvertToInternalUnits(800, UnitTypeId.Millimeters); //смещение созданного окна от уровня
                    window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(windowLevelOffset);
                    ts.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Info", "Не найден типоразмер");
            }
        }

        public void AddDoor(Document doc, Level level, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2032 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();
            if (doorType != null)
            {
                using (var ts = new Transaction(doc, "create doors"))
                {
                    ts.Start();

                    LocationCurve hostCurve = wall.Location as LocationCurve;
                    XYZ midPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) / 2;

                    if (!doorType.IsActive)
                    {
                        doorType.Activate();
                    }
                    doc.Create.NewFamilyInstance(midPoint, doorType, wall, level, StructuralType.NonStructural);

                    ts.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Info", "Не найден типоразмер");
            }
        }

        public List<Level> GetLevelList(Document doc)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();
            return listLevel;
        }
        public Level GetLevelByName(List<Level> levelList, string levelName)
        {
            return levelList
                .Where(x => x.Name.Equals(levelName))
                .FirstOrDefault();

        }
        public List<Wall> CreateWall(Document doc, List<XYZ> points, Level bottomLevel, Level topLevel)
        {

            using (var ts = new Transaction(doc, "Create walls"))
            {
                List<Wall> walls = new List<Wall>();
                ts.Start();
                for (int i = 0; i < 4; i++)
                {
                    Line line = Line.CreateBound(points[i], points[i + 1]);
                    Wall wall = Wall.Create(doc, line, bottomLevel.Id, false);
                    walls.Add(wall);
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(topLevel.Id);

                }
                ts.Commit();
                return walls;
            }
        }
        public List<XYZ> GetPointsByWidthAndDepth(double width, double depth)
        {
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));
            return points;
        }
    }
}
