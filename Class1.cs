using System.Globalization;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace NHXClipBlocks
{
    public class Commands
    {
        [CommandMethod("NHXClipBlocks", CommandFlags.Session)]
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
            using (var docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var pickedPline = (Polyline)tr.GetObject(pickResult.ObjectId, OpenMode.ForRead);
                targetLayer = pickedPline.Layer;
                tr.Commit();
            }

            ed.WriteMessage($"\nTarget rectangle layer: {targetLayer}");

            // Step 3: Separate the selection into rectangles (on targetLayer) and block references
            var blockRefIds = new List<ObjectId>();
            var rectangleIds = new List<ObjectId>();

            using (var docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;

                    Entity ent = (Entity)tr.GetObject(selObj.ObjectId, OpenMode.ForRead);

                    if (ent is BlockReference)
                    {
                        blockRefIds.Add(selObj.ObjectId);
                    }
                    else if (ent is Polyline pline
                             && pline.Closed
                             && pline.NumberOfVertices == 4
                             && string.Equals(pline.Layer, targetLayer, StringComparison.OrdinalIgnoreCase))
                    {
                        rectangleIds.Add(selObj.ObjectId);
                    }
                }

                tr.Commit();
            }

            if (blockRefIds.Count == 0)
            {
                ed.WriteMessage("\nNo block references found in selection.");
                return;
            }

            if (rectangleIds.Count == 0)
            {
                ed.WriteMessage("\nNo rectangles found on the target layer.");
                return;
            }

            // Step 4: Match each block to its rectangles, duplicate blocks as needed,
            //         and collect (blockHandle, rectVertices) pairs for XCLIP.
            // Each pair = one block ref that needs to be clipped to one rectangle.
            var clipJobs = new List<(string blockHandle, Point2d[] vertices)>();

            using (var docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var modelSpace = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                foreach (ObjectId blkId in blockRefIds)
                {
                    var blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);

                    Extents3d blkExtents;
                    try { blkExtents = blkRef.GeometricExtents; }
                    catch { continue; }

                    // Find rectangles whose center is inside this block's extents
                    var matchedRects = new List<ObjectId>();
                    foreach (ObjectId rectId in rectangleIds)
                    {
                        var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);

                        Extents3d rectExtents;
                        try { rectExtents = rect.GeometricExtents; }
                        catch { continue; }

                        var center = new Point3d(
                            (rectExtents.MinPoint.X + rectExtents.MaxPoint.X) / 2.0,
                            (rectExtents.MinPoint.Y + rectExtents.MaxPoint.Y) / 2.0,
                            0);

                        if (center.X >= blkExtents.MinPoint.X && center.X <= blkExtents.MaxPoint.X &&
                            center.Y >= blkExtents.MinPoint.Y && center.Y <= blkExtents.MaxPoint.Y)
                        {
                            matchedRects.Add(rectId);
                        }
                    }

                    if (matchedRects.Count == 0)
                    {
                        ed.WriteMessage($"\n{blkRef.Name} - 0 rectangles (skipped)");
                        continue;
                    }

                    ed.WriteMessage($"\n{blkRef.Name} - {matchedRects.Count} rectangle{(matchedRects.Count == 1 ? "" : "s")}");

                    // First rectangle -> clip the original block ref
                    var firstVerts = GetRectWcsVertices(tr, matchedRects[0]);
                    clipJobs.Add((blkRef.Handle.ToString(), firstVerts));

                    // Additional rectangles -> duplicate block ref, clip each copy
                    for (int i = 1; i < matchedRects.Count; i++)
                    {
                        var copy = new BlockReference(blkRef.Position, blkRef.BlockTableRecord)
                        {
                            Layer = blkRef.Layer,
                            ScaleFactors = blkRef.ScaleFactors,
                            Rotation = blkRef.Rotation,
                            Normal = blkRef.Normal,
                        };

                        modelSpace.AppendEntity(copy);
                        tr.AddNewlyCreatedDBObject(copy, true);

                        var verts = GetRectWcsVertices(tr, matchedRects[i]);
                        clipJobs.Add((copy.Handle.ToString(), verts));
                    }
                }

                tr.Commit();
            }

            // Step 5: Build a script string that runs _XCLIP for each block.
            // The XCLIP command sequence is:
            //   _XCLIP <select block> <enter> _New _Polygonal <pt1> <pt2> <pt3> <pt4> <enter>
            // We select each block by handle using (handent "HANDLE") LISP expression.
            if (clipJobs.Count > 0)
            {
                var script = new System.Text.StringBuilder();
                foreach (var (handle, verts) in clipJobs)
                {
                    // _XCLIP -> select via LISP handent -> enter to end selection
                    // -> _New -> _Polygonal -> 4 points -> enter to close
                    script.Append($"(command \"_.XCLIP\" (handent \"{handle}\") \"\" \"_New\" \"_Polygonal\" ");
                    foreach (var pt in verts)
                    {
                        string x = pt.X.ToString("F6", CultureInfo.InvariantCulture);
                        string y = pt.Y.ToString("F6", CultureInfo.InvariantCulture);
                        script.Append($"\"{x},{y}\" ");
                    }
                    script.Append("\"\")\n");
                }

                doc.SendStringToExecute(script.ToString(), true, false, false);
            }

            ed.WriteMessage("\nDone.\n");
        }

        /// <summary>
        /// Gets the 4 WCS vertices of a rectangle polyline.
        /// The _XCLIP command expects points in WCS.
        /// </summary>
        private static Point2d[] GetRectWcsVertices(Transaction tr, ObjectId rectId)
        {
            var rect = (Polyline)tr.GetObject(rectId, OpenMode.ForRead);

            var pts = new Point2d[rect.NumberOfVertices];
            for (int i = 0; i < rect.NumberOfVertices; i++)
            {
                Point3d wPt = rect.GetPoint3dAt(i);
                pts[i] = new Point2d(wPt.X, wPt.Y);
            }
            return pts;
        }
    }
}
