using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace NHXClipBlocks
{
    public class Commands
    {
        [CommandMethod("NHXClipBlocks")]
        public void NhxClipBlocks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Step 1: Ask user to select objects (rectangles and blocks)
            var selOpts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect rectangles and blocks: "
            };

            PromptSelectionResult selResult = ed.GetSelection(selOpts);
            if (selResult.Status != PromptStatus.OK)
                return;

            SelectionSet selSet = selResult.Value;

            // Step 2: Ask user to pick a single rectangle to determine the target layer
            var pickOpts = new PromptEntityOptions("\nSelect a rectangle to identify the clipping layer: ");
            pickOpts.SetRejectMessage("\nMust be a polyline.");
            pickOpts.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult pickResult = ed.GetEntity(pickOpts);
            if (pickResult.Status != PromptStatus.OK)
                return;

            string targetLayer;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var pickedPline = (Polyline)tr.GetObject(pickResult.ObjectId, OpenMode.ForRead);
                targetLayer = pickedPline.Layer;
                tr.Commit();
            }

            ed.WriteMessage($"\nTarget rectangle layer: {targetLayer}");

            // Step 3: Separate the selection into rectangles (on targetLayer) and block references
            var blockRefs = new List<ObjectId>();
            var rectangles = new List<ObjectId>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;

                    Entity ent = (Entity)tr.GetObject(selObj.ObjectId, OpenMode.ForRead);

                    if (ent is BlockReference)
                    {
                        blockRefs.Add(selObj.ObjectId);
                    }
                    else if (ent is Polyline pline
                             && pline.Closed
                             && pline.NumberOfVertices == 4
                             && string.Equals(pline.Layer, targetLayer, StringComparison.OrdinalIgnoreCase))
                    {
                        rectangles.Add(selObj.ObjectId);
                    }
                }

                tr.Commit();
            }

            if (blockRefs.Count == 0)
            {
                ed.WriteMessage("\nNo block references found in selection.");
                return;
            }

            if (rectangles.Count == 0)
            {
                ed.WriteMessage("\nNo rectangles found on the target layer.");
                return;
            }

            // Step 4: For each block, count how many rectangles are "on" it
            //         A rectangle is on a block if its center is inside the block's geometric extents.
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId blkId in blockRefs)
                {
                    var blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                    Extents3d blkExtents;
                    try
                    {
                        blkExtents = blkRef.GeometricExtents;
                    }
                    catch
                    {
                        // Block has no valid extents, skip
                        continue;
                    }

                    string blockName = blkRef.Name;

                    int count = 0;
                    foreach (ObjectId rectId in rectangles)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);

                        Extents3d rectExtents;
                        try
                        {
                            rectExtents = rect.GeometricExtents;
                        }
                        catch
                        {
                            continue;
                        }

                        // Use the center of the rectangle to test containment
                        var center = new Autodesk.AutoCAD.Geometry.Point3d(
                            (rectExtents.MinPoint.X + rectExtents.MaxPoint.X) / 2.0,
                            (rectExtents.MinPoint.Y + rectExtents.MaxPoint.Y) / 2.0,
                            0);

                        if (center.X >= blkExtents.MinPoint.X && center.X <= blkExtents.MaxPoint.X &&
                            center.Y >= blkExtents.MinPoint.Y && center.Y <= blkExtents.MaxPoint.Y)
                        {
                            count++;
                        }
                    }

                    if (count > 0)
                    {
                        ed.WriteMessage($"\n{blockName} - {count} rectangle{(count == 1 ? "" : "s")}");
                    }
                    else
                    {
                        ed.WriteMessage($"\n{blockName} - 0 rectangles");
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage("\n");
        }
    }
}
