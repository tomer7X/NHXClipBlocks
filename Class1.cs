using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Linq;

namespace NHXClipBlocks
{
    public class Commands
    {
        private const int NetColumns = 10;

        [CommandMethod("NHXClipBlocks")]
        public void NhxClipBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            // Prompt inputs
            PromptSelectionResult selResult = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\nSelect rectangles and blocks: " });
            if (selResult.Status != PromptStatus.OK) return;
            SelectionSet selSet = selResult.Value;

            var pickOpts = new PromptEntityOptions("\nSelect a rectangle to identify the clipping layer: ");
            pickOpts.SetRejectMessage("\nMust be a polyline.");
            pickOpts.AddAllowedClass(typeof(Polyline), true);
            PromptEntityResult pickResult = ed.GetEntity(pickOpts);
            if (pickResult.Status != PromptStatus.OK) return;

            var netOpts = new PromptEntityOptions("\nSelect the NET block: ");
            netOpts.SetRejectMessage("\nMust be a block reference.");
            netOpts.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult netResult = ed.GetEntity(netOpts);
            if (netResult.Status != PromptStatus.OK) return;

            // Ask orientation (Vertical default)
            var orientOpts = new PromptKeywordOptions("\nOrientation [Vertical/Horizontal] <Vertical>: ", "Vertical Horizontal");
            orientOpts.AllowNone = true;
            PromptResult orientRes = ed.GetKeywords(orientOpts);
            bool isHorizontal = orientRes.Status == PromptStatus.OK && orientRes.StringResult == "Horizontal";

            // Build work list
            if (!BuildWorkList(db, selSet, pickResult.ObjectId, netResult.ObjectId, ed, out string targetLayer, out Extents3d netExtents, out var blockData))
                return;

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

            // list of (blkCopyId, rectCopyId, rectExt, cell)
            var clipWork = CopyToNet(db, blockData, netMinX, netMaxY, cellWidth, cellHeight, isHorizontal);
            ed.WriteMessage($"\nPlaced {blockData.Count} group(s), {clipWork.Count} block copies total.");

            int applied = ApplyXClips(db, clipWork, ed);
            ed.WriteMessage($"\nDone. Applied {applied}/{clipWork.Count} XCLIP(s).\n");

            for (int i = 0; i < clipWork.Count; i++)
            {
                var (blkCopyId, rectCopyId, rectExt, cell) = clipWork[i];
                ed.WriteMessage($"\n  {i + 1}. Cell {cell}, Block {blkCopyId}, Rect {rectCopyId}, Extents: ({rectExt.MinPoint.X}, {rectExt.MinPoint.Y}) to ({rectExt.MaxPoint.X}, {rectExt.MaxPoint.Y})");
            }

            

            // Collect block references grouped by NET cell so the caller can adjust positions
            var cells = GetBlocksByCell(db, clipWork, isHorizontal);
            for (int i = 0; i < cells.Length; i++)
            {
                // arrange as a list: move each item (index k) to sit below previous (k-1)
                for (int k = 1; k < cells[i].Count; k++)
                {
                    AttachRectangles(db, clipWork, i, k, k - 1, isHorizontal);
                }
            }
        }


        private bool BuildWorkList(
            Database db,
            SelectionSet selSet,
            ObjectId pickPlineId,
            ObjectId netBlockId,
            Editor ed,
            out string targetLayer,
            out Extents3d netExtents,
            out List<(ObjectId blkId, double blkX, List<ObjectId> rects)> blockData)
        {
            targetLayer = null;
            netExtents = default;
            blockData = new List<(ObjectId, double, List<ObjectId>)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var pickedPline = (Polyline)tr.GetObject(pickPlineId, OpenMode.ForRead);
                targetLayer = pickedPline.Layer;
                ed.WriteMessage($"\nTarget rectangle layer: {targetLayer}");

                var netRef = (BlockReference)tr.GetObject(netBlockId, OpenMode.ForRead);
                try { netExtents = netRef.GeometricExtents; }
                catch
                {
                    ed.WriteMessage("\nNET block has no valid extents.");
                    tr.Commit();
                    return false;
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
                    return false;
                }
                if (rectangleIds.Count == 0)
                {
                    ed.WriteMessage("\nNo rectangles found on the target layer.");
                    tr.Commit();
                    return false;
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

            return true;
        }

        private List<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)> CopyToNet(
            Database db,
            List<(ObjectId blkId, double blkX, List<ObjectId> rects)> blockData,
            double netMinX,
            double netMaxY,
            double cellWidth,
            double cellHeight,
            bool isHorizontal)
        {
            var clipWork = new List<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

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

                    foreach (ObjectId rectId in rects)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);
                        Extents3d origRectExt;
                        try { origRectExt = rect.GeometricExtents; }
                        catch { continue; }

                        var shiftedRectExt = new Extents3d(
                            new Point3d(origRectExt.MinPoint.X + dx, origRectExt.MinPoint.Y + dy, 0),
                            new Point3d(origRectExt.MaxPoint.X + dx, origRectExt.MaxPoint.Y + dy, 0));

                        var rectCopy = (Polyline)rect.Clone();
                        rectCopy.TransformBy(displacement);
                        modelSpace.AppendEntity(rectCopy);
                        tr.AddNewlyCreatedDBObject(rectCopy, true);

                        var blkCopy = (BlockReference)blkRef.Clone();
                        blkCopy.TransformBy(displacement);
                        modelSpace.AppendEntity(blkCopy);
                        tr.AddNewlyCreatedDBObject(blkCopy, true);

                        foreach (ObjectId attId in blkRef.AttributeCollection)
                        {
                            var attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                            var attCopy = (AttributeReference)attRef.Clone();
                            attCopy.TransformBy(displacement);
                            blkCopy.AttributeCollection.AppendAttribute(attCopy);
                            tr.AddNewlyCreatedDBObject(attCopy, true);
                        }

                        int cell = idx; // zero-based cell index (row*NetColumns + col)
                        clipWork.Add((blkCopy.ObjectId, rectCopy.ObjectId, shiftedRectExt, cell));
                    }
                }

                tr.Commit();
            }

            // Now sort the clipWork within each cell according to orientation so later grouping preserves order
            var sorted = new List<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)>();
            var groups = clipWork.GroupBy(x => x.cell).OrderBy(g => g.Key);
            foreach (var g in groups)
            {
                IEnumerable<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)> ordered;
                if (isHorizontal)
                {
                    ordered = g.OrderBy(x => x.rectExt.MinPoint.X).ThenByDescending(x => x.rectExt.MaxPoint.Y);
                }
                else
                {
                    ordered = g.OrderByDescending(x => x.rectExt.MaxPoint.Y).ThenBy(x => x.rectExt.MinPoint.X);
                }
                sorted.AddRange(ordered);
            }

            return sorted;
        }

        private int ApplyXClips(Database db, List<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)> clipWork, Editor ed)
        {
            int clipCount = 0;
            foreach (var (blkCopyId, _, rectExt, _) in clipWork)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        var blkRef = (BlockReference)tr.GetObject(blkCopyId, OpenMode.ForWrite);

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

            return clipCount;
        }

        private List<ObjectId>[] GetBlocksByCell(
    Database db,
    List<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)> clipWork,
    bool isHorizontal)
        {
            int maxCell = -1;
            foreach (var item in clipWork)
            {
                if (item.cell > maxCell)
                    maxCell = item.cell;
            }

            int cellCount = Math.Max(0, maxCell + 1);
            var result = new List<ObjectId>[cellCount];
            for (int i = 0; i < cellCount; i++)
                result[i] = new List<ObjectId>();

            foreach (var item in clipWork)
            {
                if (item.cell >= 0 && item.cell < cellCount)
                    result[item.cell].Add(item.blkCopyId);
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var sortKeys = new Dictionary<ObjectId, double>();

                foreach (var item in clipWork)
                {
                    if (!sortKeys.ContainsKey(item.blkCopyId))
                    {
                        var br = tr.GetObject(item.blkCopyId, OpenMode.ForRead) as BlockReference;
                        if (br != null)
                        {
                            var ext = br.GeometricExtents;
                            sortKeys[item.blkCopyId] = isHorizontal
                                ? ext.MinPoint.X
                                : ext.MaxPoint.Y;
                        }
                    }
                }

                for (int i = 0; i < cellCount; i++)
                {
                    if (isHorizontal)
                    {
                        result[i].Sort((a, b) => sortKeys[a].CompareTo(sortKeys[b]));
                    }
                    else
                    {
                        result[i].Sort((a, b) => sortKeys[b].CompareTo(sortKeys[a]));
                    }
                }

                tr.Commit();
            }

            return result;
        }

        private void AttachRectangles(Database db, List<(ObjectId blkCopyId, ObjectId rectCopyId, Extents3d rectExt, int cell)> clipWork, int cellIndex, int sourceIndex, int destIndex, bool isHorizontal)
        {
            // Generic: move the rectangle at sourceIndex so it is positioned relative to destIndex
            // destIndex rectangle is assumed to be above the source rectangle. We will move the source
            // so its top (maxY) is just below the dest's bottom (minY) with the specified margin (5.0).
            if (clipWork == null)
                return;

            // collect entries for the requested cell preserving clipWork order
            var entries = new List<(ObjectId blkId, ObjectId rectId, Extents3d rectExt)>();
            foreach (var item in clipWork)
            {
                if (item.cell == cellIndex)
                    entries.Add((item.blkCopyId, item.rectCopyId, item.rectExt));
            }

            if (sourceIndex < 0 || destIndex < 0 || sourceIndex >= entries.Count || destIndex >= entries.Count)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var src = entries[sourceIndex];
                    var dst = entries[destIndex];

                    var srcRef = tr.GetObject(src.blkId, OpenMode.ForWrite) as BlockReference;
                    var srcRect = tr.GetObject(src.rectId, OpenMode.ForWrite) as Polyline;
                    var dstRect = tr.GetObject(dst.rectId, OpenMode.ForRead) as Polyline;
                    if (srcRef == null || srcRect == null || dstRect == null)
                    {
                        tr.Commit();
                        return;
                    }

                    // Determine whether rectangles are aligned by X (vertical stack) or by Y (horizontal list)
                    double srcCenterX = (src.rectExt.MinPoint.X + src.rectExt.MaxPoint.X) / 2.0;
                    double srcCenterY = (src.rectExt.MinPoint.Y + src.rectExt.MaxPoint.Y) / 2.0;
                    double dstCenterX = (dst.rectExt.MinPoint.X + dst.rectExt.MaxPoint.X) / 2.0;
                    double dstCenterY = (dst.rectExt.MinPoint.Y + dst.rectExt.MaxPoint.Y) / 2.0;

                    double dx = System.Math.Abs(srcCenterX - dstCenterX);
                    double dy = System.Math.Abs(srcCenterY - dstCenterY);

                    double margin = 5.0;
                    double offsetX = 0.0;
                    double offsetY = 0.0;

                    if (!isHorizontal && dx < dy)
                    {
                        // X coordinates are approximately equal -> vertical stacking (move in Y)
                        double srcMaxY = src.rectExt.MaxPoint.Y;
                        double dstMinY = dst.rectExt.MinPoint.Y;
                        offsetY = (dstMinY - margin) - srcMaxY;
                    }
                    else
                    {
                        // Y coordinates are approximately equal -> horizontal list (move in X)
                        double srcMinX = src.rectExt.MinPoint.X;
                        double dstMaxX = dst.rectExt.MaxPoint.X;
                        offsetX = (dstMaxX + margin) - srcMinX;
                    }

                    var disp = Matrix3d.Displacement(new Vector3d(offsetX, offsetY, 0));

                    // Move block reference and its attributes
                    srcRef.TransformBy(disp);
                    foreach (ObjectId attId in srcRef.AttributeCollection)
                    {
                        var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                        if (attRef != null)
                            attRef.TransformBy(disp);
                    }

                    // Also move the copied clipping rectangle so it stays with the block
                    try
                    {
                        if (srcRect != null && !srcRect.IsErased)
                            srcRect.TransformBy(disp);
                    }
                    catch { }

                    // Update the stored extents in clipWork so subsequent moves use the new position
                    for (int ci = 0; ci < clipWork.Count; ci++)
                    {
                        var it = clipWork[ci];
                        if (it.rectCopyId == src.rectId && it.blkCopyId == src.blkId && it.cell == cellIndex)
                        {
                            var newExt = new Extents3d(
                                new Point3d(it.rectExt.MinPoint.X + offsetX, it.rectExt.MinPoint.Y + offsetY, 0),
                                new Point3d(it.rectExt.MaxPoint.X + offsetX, it.rectExt.MaxPoint.Y + offsetY, 0));
                            clipWork[ci] = (it.blkCopyId, it.rectCopyId, newExt, it.cell);
                            break;
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }
        }
    }
}
