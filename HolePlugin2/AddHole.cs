using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HolePlugin2
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;//обращение к ар документу
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();//выделение ов документа
            if (ovDoc == null)
            {
                TaskDialog.Show("Error", "Файл ОВ не найден");
                return Result.Cancelled;
            }
            //проверка на наличие семейтсва с отверстиями
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстия"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Error", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }
            //поиск воздуховодов
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();
            //Поиск труб(задание)
            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();
            //Поиск #D виде для  ReferenceIntersector
            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)//исключение шаблонов
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Error", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transaction1 = new Transaction(arDoc);
            transaction1.Start("Активация символа");
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            transaction1.Commit();


            //перебор воздуховодов и вставка отверстия в модель
            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");
            DuctHoles(arDoc, referenceIntersector, ducts, familySymbol);
            PipeHoles(arDoc, referenceIntersector, pipes, familySymbol);

            transaction.Commit();

            return Result.Succeeded;
        }

        private void PipeHoles(Document arDoc, ReferenceIntersector referenceIntersector, List<Pipe> pipes, FamilySymbol familySymbol)
        {
            foreach (Pipe p in pipes)
            {

                Line curve = (p.Location as LocationCurve).Curve as Line;

                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                //поиск набора всех перечечений
                List<ReferenceWithContext> intersection = referenceIntersector.Find(point, direction)
                     .Where(x => x.Proximity <= curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();

                //перебор набора пересечений и определение точки вставки
                foreach (ReferenceWithContext refer in intersection)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(p.Diameter);
                    height.Set(p.Diameter);
                }
            }
        }
        public void DuctHoles(Document arDoc, ReferenceIntersector referenceIntersector, List<Duct> ducts, FamilySymbol familySymbol)
        {
            foreach (Duct d in ducts)
            {

                Line curve = (d.Location as LocationCurve).Curve as Line;

                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                //поиск набора всех перечечений
                List<ReferenceWithContext> intersection = referenceIntersector.Find(point, direction)
                     .Where(x => x.Proximity <= curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();

                //перебор набора пересечений и определение точки вставки
                foreach (ReferenceWithContext refer in intersection)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(d.Diameter);
                    height.Set(d.Diameter);
                }
            }

        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }

    }
}
