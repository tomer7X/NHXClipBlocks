using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace NHXClipBlocks
{
    public class Commands
    {
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

            // ── Gather data in a read transaction, then copy + clip ──
            string targetLayer;
            Extents3d netExtents;
            var blockData = new List<(ObjectId blkId, double blkX, List<ObjectId> rects)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var pickedPline = (Polyline)tr.GetObject(pickResult.ObjectId, OpenMode.ForRead);
                targetLayer = pickedPline.Layer;
                ed.WriteMessage($"\nTarget rectangle layer: {targetLayer}");

                var netRef = (BlockReference)tr.GetObject(netResult.ObjectId, OpenMode.ForRead);
                try { netExtents = netRef.GeometricExtents; }
                catch
                {
                    ed.WriteMessage("\nNET block has no valid extents.");
                    tr.Commit();
                    return;
                }

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

                tr.Commit();
            }

            if (blockData.Count == 0)
            {
                ed.WriteMessage("\nNo blocks matched any rectangles.");
                return;
            }

            blockData.Sort((a, b) => a.blkX.CompareTo(b.blkX));

            double netMinX = netExtents.MinPoint.X;
            double netMaxX = netExtents.MaxPoint.X;
            double netMaxY = netExtents.MaxPoint.Y;
            double cellWidth = (netMaxX - netMinX) / NetColumns;
            double cellHeight = cellWidth / 2.0;

            // ── Copy blocks + rectangles into NET, one block copy per rectangle ──
            // Store (blockCopyId, rectExtents) pairs for XCLIP later.
            var clipWork = new List<(ObjectId blkCopyId, Extents3d rectExt)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var modelSpace = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                for (int idx = 0; idx < blockData.Count; idx++)
                {
                    var (blkId, _, rects) = blockData[idx];
                    var blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                    Extents3d groupExtents = blkRef.GeometricExtents;
                    foreach (ObjectId rectId in rects)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);
                        try { groupExtents.AddExtents(rect.GeometricExtents); }
                        catch { }
                    }

                    double groupCenterX = (groupExtents.MinPoint.X + groupExtents.MaxPoint.X) / 2.0;
                    double groupCenterY = (groupExtents.MinPoint.Y + groupExtents.MaxPoint.Y) / 2.0;

                    int col = idx % NetColumns;
                    int row = idx / NetColumns;
                    double cellCenterX = netMinX + col * cellWidth + cellWidth / 2.0;
                    double cellCenterY = netMaxY - row * cellHeight - cellHeight / 2.0;

                    double dx = cellCenterX - groupCenterX;
                    double dy = cellCenterY - groupCenterY;
                    var displacement = Matrix3d.Displacement(new Vector3d(dx, dy, 0));

                    // Read each rectangle's extents (shifted) and create one block copy per rect
                    foreach (ObjectId rectId in rects)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);
                        Extents3d origRectExt;
                        try { origRectExt = rect.GeometricExtents; }
                        catch { continue; }

                        // Shifted rectangle extents in WCS
                        var shiftedRectExt = new Extents3d(
                            new Point3d(origRectExt.MinPoint.X + dx, origRectExt.MinPoint.Y + dy, 0),
                            new Point3d(origRectExt.MaxPoint.X + dx, origRectExt.MaxPoint.Y + dy, 0));

                        // Copy the rectangle
                        var rectCopy = (Polyline)rect.Clone();
                        rectCopy.TransformBy(displacement);
                        modelSpace.AppendEntity(rectCopy);
                        tr.AddNewlyCreatedDBObject(rectCopy, true);

                        // Clone the block (full deep copy preserves all internal state)
                        var blkCopy = (BlockReference)blkRef.Clone();
                        blkCopy.TransformBy(displacement);
                        modelSpace.AppendEntity(blkCopy);
                        tr.AddNewlyCreatedDBObject(blkCopy, true);

                        // Deep-copy attribute references so they appear on the clone
                        foreach (ObjectId attId in blkRef.AttributeCollection)
                        {
                            var attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                            var attCopy = (AttributeReference)attRef.Clone();
                            attCopy.TransformBy(displacement);
                            blkCopy.AttributeCollection.AppendAttribute(attCopy);
                            tr.AddNewlyCreatedDBObject(attCopy, true);
                        }

                        clipWork.Add((blkCopy.ObjectId, shiftedRectExt));
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nPlaced {blockData.Count} group(s), {clipWork.Count} block copies total.");

            // ── Apply XCLIP to each block copy, one transaction per clip ──
            int clipCount = 0;
            foreach (var (blkCopyId, rectExt) in clipWork)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        var blkRef = (BlockReference)tr.GetObject(blkCopyId, OpenMode.ForWrite);

                        // Transform WCS boundary corners into block-local coordinates
                        Matrix3d inv = blkRef.BlockTransform.Inverse();

                        Point3d p1 = new Point3d(rectExt.MinPoint.X, rectExt.MinPoint.Y, 0).TransformBy(inv);
                        Point3d p2 = new Point3d(rectExt.MaxPoint.X, rectExt.MinPoint.Y, 0).TransformBy(inv);
                        Point3d p3 = new Point3d(rectExt.MaxPoint.X, rectExt.MaxPoint.Y, 0).TransformBy(inv);
                        Point3d p4 = new Point3d(rectExt.MinPoint.X, rectExt.MaxPoint.Y, 0).TransformBy(inv);

                        var clipBoundary = new Point2dCollection
                        {
                            new Point2d(p1.X, p1.Y),
                            new Point2d(p2.X, p2.Y),
                            new Point2d(p3.X, p3.Y),
                            new Point2d(p4.X, p4.Y),
                        };

                        var spatialDef = new SpatialFilterDefinition(
                            clipBoundary,
                            Vector3d.ZAxis,
                            0.0,
                            double.PositiveInfinity,
                            double.NegativeInfinity,
                            true);

                        var spatialFilter = new SpatialFilter { Definition = spatialDef };

                        if (blkRef.ExtensionDictionary == ObjectId.Null)
                            blkRef.CreateExtensionDictionary();

                        var extDict = (DBDictionary)tr.GetObject(blkRef.ExtensionDictionary, OpenMode.ForWrite);

                        DBDictionary filterDict;
                        if (extDict.Contains("ACAD_FILTER"))
                        {
                            filterDict = (DBDictionary)tr.GetObject(extDict.GetAt("ACAD_FILTER"), OpenMode.ForWrite);
                        }
                        else
                        {
                            filterDict = new DBDictionary();
                            extDict.SetAt("ACAD_FILTER", filterDict);
                            tr.AddNewlyCreatedDBObject(filterDict, true);
                        }

                        if (filterDict.Contains("SPATIAL"))
                        {
                            var existing = (SpatialFilter)tr.GetObject(filterDict.GetAt("SPATIAL"), OpenMode.ForWrite);
                            existing.Definition = spatialDef;
                            spatialFilter.Dispose();
                        }
                        else
                        {
                            filterDict.SetAt("SPATIAL", spatialFilter);
                            tr.AddNewlyCreatedDBObject(spatialFilter, true);
                        }

                        tr.Commit();
                        clipCount++;
                    }
                    catch (System.Exception ex)
                    {
                        tr.Abort();
                        ed.WriteMessage($"\n  XCLIP failed: {ex.Message}");
                    }
                }
            }

            ed.WriteMessage($"\nDone. Applied {clipCount}/{clipWork.Count} XCLIP(s).\n");
        }
    }
}
