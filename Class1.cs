using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace NHXClipBlocks
{
    public class Commands
    {
        private const double Margin = 5.0;
        private const int NetColumns = 10;

        [CommandMethod("NHXClipBlocks")]
        public void NhxClipBlocks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // ── Step 1: Select rectangles and blocks ──
            PromptSelectionResult selResult = ed.GetSelection(new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect rectangles and blocks: "
            });
            if (selResult.Status != PromptStatus.OK) return;
            SelectionSet selSet = selResult.Value;

            // ── Step 2: Pick a rectangle to identify the clipping layer ──
            var pickOpts = new PromptEntityOptions("\nSelect a rectangle to identify the clipping layer: ");
            pickOpts.SetRejectMessage("\nMust be a polyline.");
            pickOpts.AddAllowedClass(typeof(Polyline), true);
            PromptEntityResult pickResult = ed.GetEntity(pickOpts);
            if (pickResult.Status != PromptStatus.OK) return;

            // ── Step 3: Pick the NET block ──
            var netOpts = new PromptEntityOptions("\nSelect the NET block: ");
            netOpts.SetRejectMessage("\nMust be a block reference.");
            netOpts.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult netResult = ed.GetEntity(netOpts);
            if (netResult.Status != PromptStatus.OK) return;

            // ── All work in a single transaction ──
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Target layer from picked rectangle
                var pickedPline = (Polyline)tr.GetObject(pickResult.ObjectId, OpenMode.ForRead);
                string targetLayer = pickedPline.Layer;
                ed.WriteMessage($"\nTarget rectangle layer: {targetLayer}");

                // NET block extents
                var netRef = (BlockReference)tr.GetObject(netResult.ObjectId, OpenMode.ForRead);
                Extents3d netExtents;
                try { netExtents = netRef.GeometricExtents; }
                catch
                {
                    ed.WriteMessage("\nNET block has no valid extents.");
                    tr.Commit();
                    return;
                }

                // Classify selected objects into blocks and rectangles
                var blockRefIds = new List<ObjectId>();
                var rectangleIds = new List<ObjectId>();

                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;
                    Entity ent = (Entity)tr.GetObject(selObj.ObjectId, OpenMode.ForRead);

                    if (ent is BlockReference)
                        blockRefIds.Add(selObj.ObjectId);
                    else if (ent is Polyline pline
                             && pline.Closed
                             && pline.NumberOfVertices == 4
                             && string.Equals(pline.Layer, targetLayer, StringComparison.OrdinalIgnoreCase))
                        rectangleIds.Add(selObj.ObjectId);
                }

                if (blockRefIds.Count == 0)
                {
                    ed.WriteMessage("\nNo block references found in selection.");
                    tr.Commit();
                    return;
                }
                if (rectangleIds.Count == 0)
                {
                    ed.WriteMessage("\nNo rectangles found on the target layer.");
                    tr.Commit();
                    return;
                }

                // NET cell dimensions
                double netMinX = netExtents.MinPoint.X;
                double netMaxX = netExtents.MaxPoint.X;
                double netMaxY = netExtents.MaxPoint.Y;
                double cellWidth = (netMaxX - netMinX) / NetColumns;
                double cellHeight = cellWidth / 2.0;

                // Match each block to rectangles whose center falls inside the block's extents
                var blockData = new List<(ObjectId blkId, double blkX, List<ObjectId> rects)>();

                foreach (ObjectId blkId in blockRefIds)
                {
                    var blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                    Extents3d blkExtents;
                    try { blkExtents = blkRef.GeometricExtents; }
                    catch { continue; }

                    var matchedRects = new List<ObjectId>();
                    foreach (ObjectId rectId in rectangleIds)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);
                        Extents3d rectExtents;
                        try { rectExtents = rect.GeometricExtents; }
                        catch { continue; }

                        double cx = (rectExtents.MinPoint.X + rectExtents.MaxPoint.X) / 2.0;
                        double cy = (rectExtents.MinPoint.Y + rectExtents.MaxPoint.Y) / 2.0;

                        if (cx >= blkExtents.MinPoint.X && cx <= blkExtents.MaxPoint.X &&
                            cy >= blkExtents.MinPoint.Y && cy <= blkExtents.MaxPoint.Y)
                            matchedRects.Add(rectId);
                    }

                    if (matchedRects.Count > 0)
                    {
                        blockData.Add((blkId, blkExtents.MinPoint.X, matchedRects));
                        ed.WriteMessage($"\n{blkRef.Name} - {matchedRects.Count} rectangle{(matchedRects.Count == 1 ? "" : "s")}");
                    }
                }

                if (blockData.Count == 0)
                {
                    ed.WriteMessage("\nNo blocks matched any rectangles.");
                    tr.Commit();
                    return;
                }

                // Sort blocks left-to-right
                blockData.Sort((a, b) => a.blkX.CompareTo(b.blkX));

                // Place copies into NET cells
                var modelSpace = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                int totalCopies = 0;

                for (int idx = 0; idx < blockData.Count; idx++)
                {
                    var (blkId, _, rects) = blockData[idx];
                    var blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                    // Compute bounding box of the whole group (block + all its rectangles)
                    Extents3d groupExtents = blkRef.GeometricExtents;
                    foreach (ObjectId rectId in rects)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);
                        try { groupExtents.AddExtents(rect.GeometricExtents); }
                        catch { }
                    }

                    double groupCenterX = (groupExtents.MinPoint.X + groupExtents.MaxPoint.X) / 2.0;
                    double groupCenterY = (groupExtents.MinPoint.Y + groupExtents.MaxPoint.Y) / 2.0;

                    // Which NET cell does this block go into
                    int col = idx % NetColumns;
                    int row = idx / NetColumns;
                    double cellCenterX = netMinX + col * cellWidth + cellWidth / 2.0;
                    double cellCenterY = netMaxY - row * cellHeight - cellHeight / 2.0;

                    // Displacement: move group center to cell center
                    double dx = cellCenterX - groupCenterX;
                    double dy = cellCenterY - groupCenterY;

                    // Copy each rectangle (shifted)
                    foreach (ObjectId rectId in rects)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);
                        var rectCopy = (Polyline)rect.Clone();
                        rectCopy.TransformBy(Matrix3d.Displacement(new Vector3d(dx, dy, 0)));
                        modelSpace.AppendEntity(rectCopy);
                        tr.AddNewlyCreatedDBObject(rectCopy, true);
                    }

                    // Copy the block N times (one per rectangle), all shifted
                    for (int i = 0; i < rects.Count; i++)
                    {
                        var blkCopy = new BlockReference(
                            new Point3d(blkRef.Position.X + dx, blkRef.Position.Y + dy, blkRef.Position.Z),
                            blkRef.BlockTableRecord)
                        {
                            Layer = blkRef.Layer,
                            ScaleFactors = blkRef.ScaleFactors,
                            Rotation = blkRef.Rotation,
                            Normal = blkRef.Normal,
                        };
                        modelSpace.AppendEntity(blkCopy);
                        tr.AddNewlyCreatedDBObject(blkCopy, true);
                        totalCopies++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\nDone. Placed {blockData.Count} block group(s) into NET cells ({totalCopies} block copies).\n");
            }
        }
    }
}
