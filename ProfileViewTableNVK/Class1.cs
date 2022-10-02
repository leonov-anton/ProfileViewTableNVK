using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;


namespace ProfileViewTableNVK
{
    public class Class1
    {
        [CommandMethod("PVNVKTABLE")]
        static public void CreateBlockTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            CivilDocument cdoc = CivilDocument.GetCivilDocument(db);

            PromptSelectionOptions pso = new PromptSelectionOptions
            {
                MessageForAdding = "\nВыбирете виды профилей",
                SinglePickInSpace = true
            };
            PromptSelectionResult psr = ed.GetSelection(pso);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВыполнение прервано");
                return;
            }

            ObjectIdCollection pipe_nets = cdoc.GetPipeNetworkIds();

            using (var t = doc.TransactionManager.StartTransaction())
            {
                ObjectIdCollection selected_prof_view = new ObjectIdCollection();

                foreach (ObjectId objid in psr.Value.GetObjectIds())
                {
                    if (objid.ObjectClass.Name == "AeccDbGraphProfile")
                        selected_prof_view.Add(objid);
                }

                if (selected_prof_view.Count == 0)
                {
                    ed.WriteMessage("\nВиды профиля не выбраны");
                    t.Commit();
                    return;
                }

                foreach (ObjectId prof_view_id in selected_prof_view)
                {
                    ProfileView profileView = (ProfileView)t.GetObject(prof_view_id, OpenMode.ForRead);
                    Alignment profileViewAligment = (Alignment)t.GetObject(profileView.AlignmentId, OpenMode.ForRead);

                    if (false)
                    {
                        return;
                    }

                    else
                    {
                        foreach (ObjectId profId in profileViewAligment.GetProfileIds())
                        {
                            Profile profile = (Profile)t.GetObject(profId, OpenMode.ForWrite);
                            if (profile.Name.ToLower().Contains("лоток"))
                                profile.Erase();
                        }

                        List<string> overrideStylePipesName = new List<string>();
                        foreach (PipeOverride pipe in profileView.PipeOverrides)
                        {
                            if (pipe.UseOverrideStyle)
                            {
                                overrideStylePipesName.Add(pipe.PipeName);
                            }
                        }
                        
                        DBObjectCollection pipes = new DBObjectCollection();

                        foreach (ObjectId pipe_net_id in pipe_nets)
                        {
                            Network pipenet = (Network)t.GetObject(pipe_net_id, OpenMode.ForRead);
                            ObjectIdCollection pipeIds = pipenet.GetPipeIds();
                            foreach (ObjectId pipeId in pipeIds)
                            {
                                Pipe pipeObj = (Pipe)t.GetObject(pipeId, OpenMode.ForWrite);
                                if (pipeObj.GetProfileViewsDisplayingMe().Contains(prof_view_id) 
                                    && !pipeObj.NetworkName.ToLower().Contains("футляр") 
                                    && !overrideStylePipesName.Contains(pipeObj.Name))
                                {
                                    pipes.Add(pipeObj);
                                }
                            }
                        }

                        BlockErase(db, $"К2_трубы_обозначение трубы и тип изоляции-{profileViewAligment.Name}");
                        BlockErase(db, $"К2_трубы_обозначение уклона и длины-{profileViewAligment.Name}");

                        BlockTable bt = (BlockTable)t.GetObject(db.BlockTableId, OpenMode.ForWrite);

                        BlockTableRecord slopeBtr = new BlockTableRecord();
                        slopeBtr.Name = $"К2_трубы_обозначение уклона и длины-{profileViewAligment.Name}";
                        ObjectId slopeBtrId = bt.Add(slopeBtr);
                        t.AddNewlyCreatedDBObject(slopeBtr, true);

                        Point3d point1 = new Point3d(5, 0, 0);
                        Point3d point2 = new Point3d(5, 5, 0);
                        Point3d point3 = new Point3d(0, 0, 0);
                        Point3d point4 = new Point3d(0, 0, 0);

                        Line startLine = new Line(point1, point2);
                        slopeBtr.AppendEntity(startLine);
                        t.AddNewlyCreatedDBObject(startLine, true);

                        double partLen = 0;

                        List<Pipe> pipesChain = new List<Pipe>();

                        foreach (Pipe pipe in pipes)
                        {
                            if (pipe.StartPoint.X == profileViewAligment.StartPoint.X && pipe.StartPoint.Y == profileViewAligment.StartPoint.Y)
                            {
                                pipes.Remove(pipe);
                                pipesChain.Add(pipe);
                                break;
                            } 
                            else if (pipe.EndPoint.X == profileViewAligment.StartPoint.X && pipe.EndPoint.Y == profileViewAligment.StartPoint.Y)
                            {
                                ReversePipe(pipe);
                                pipes.Remove(pipe);
                                pipesChain.Add(pipe);
                                break;
                            }
                        }

                        while (pipes.Count > 0)
                        {
                            ObjectId privStructId = pipesChain.Last().EndStructureId;
                            Structure privStructObj = (Structure)t.GetObject(privStructId, OpenMode.ForRead);
                            var structPipesNames = privStructObj.GetConnectedPipeNames();

                            foreach (Pipe pipe in pipes)
                            {
                                if (structPipesNames.Contains(pipe.Name))
                                {
                                    if (pipe.EndStructureId == privStructId)
                                        ReversePipe(pipe);
                                    
                                    double station = Double.NaN;
                                    double offset = Double.NaN;

                                    profileViewAligment.StationOffset(pipe.StartPoint.X, pipe.StartPoint.Y, ref station, ref offset);

                                    if (Math.Round(pipe.Slope, 3) == Math.Round(pipesChain.Last().Slope, 3) 
                                        && (Math.Round(pipesChain.Last().EndPoint.Z - pipesChain.Last().InnerDiameterOrWidth / 2, 1) == Math.Round(pipe.StartPoint.Z - pipe.InnerDiameterOrWidth / 2, 1) 
                                        || Math.Round(pipesChain.Last().EndPoint.Z + pipesChain.Last().InnerDiameterOrWidth / 2, 1) == Math.Round(pipe.StartPoint.Z + pipe.InnerDiameterOrWidth / 2, 1)))
                                    {
                                        partLen += pipesChain.Last().Length2D;
                                    }
                                    else
                                    {
                                        partLen += pipesChain.Last().Length2D;
                                        
                                        point3 = new Point3d(5 + station, 0, 0);
                                        point4 = new Point3d(5 + station, 5, 0);

                                        CreateSlopeBlockPart(db, slopeBtr, pipesChain.Last(), point1, point2, point3, point4, partLen);

                                        point1 = new Point3d(point3.X, point3.Y, 0);
                                        point2 = new Point3d(point4.X, point4.Y, 0);

                                        partLen = 0;

                                    }
                                
                                    pipes.Remove(pipe);
                                    pipesChain.Add(pipe);
                                    break;

                                }
                            }
                        }

                        point3 = new Point3d(5 + profileViewAligment.EndingStation, 0, 0);
                        point4 = new Point3d(5 + profileViewAligment.EndingStation, 5, 0);

                        partLen += pipesChain.Last().Length2D;

                        CreateSlopeBlockPart(db, slopeBtr, pipesChain.Last(), point1, point2, point3, point4, partLen);

                        BlockTableRecord btr = (BlockTableRecord)t.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        BlockReference br = new BlockReference(new Point3d(profileView.Location.X, profileView.Location.Y - 40, profileView.Location.Z), slopeBtrId);
                        btr.AppendEntity(br);
                        t.AddNewlyCreatedDBObject(br, true);

                        BlockTableRecord typeBtr = new BlockTableRecord();
                        typeBtr.Name = $"К2_трубы_обозначение трубы и тип изоляции-{profileViewAligment.Name}";
                        ObjectId typeBtrId = bt.Add(typeBtr);
                        t.AddNewlyCreatedDBObject(typeBtr, true);

                        startLine = new Line(new Point3d(5, 0, 0), new Point3d(5, 7.5, 0));
                        typeBtr.AppendEntity(startLine);
                        t.AddNewlyCreatedDBObject(startLine, true);

                        double privStation = 0;

                        for (byte i = 1; i < pipesChain.Count ; i++)
                        {
                            if (pipesChain[i].PartSizeName != pipesChain[i - 1].PartSizeName)
                            {
                                double station = Double.NaN;
                                double offset = Double.NaN;

                                profileViewAligment.StationOffset(pipesChain[i - 1].EndPoint.X, pipesChain[i - 1].EndPoint.Y, ref station, ref offset);

                                Line nextLine = new Line(new Point3d(5 + station, 0, 0), new Point3d(5 + station, 7.5, 0));
                                typeBtr.AppendEntity(nextLine);
                                t.AddNewlyCreatedDBObject(nextLine, true);

                                CreateTypeBlockPart(db, typeBtr, pipesChain[i - 1], privStation, station);

                                privStation = station;
                            }
                        }

                        CreateTypeBlockPart(db, typeBtr, pipesChain.Last(), privStation, profileViewAligment.EndingStation);

                        br = new BlockReference(new Point3d(profileView.Location.X, profileView.Location.Y - 30, profileView.Location.Z), typeBtrId);
                        btr.AppendEntity(br);
                        t.AddNewlyCreatedDBObject(br, true);

                    }

                }
                t.Commit();
            }

        }

        static void BlockErase(Database db, string blkName)
        {

            ObjectId blkId = new ObjectId();

            using (Transaction tmpt = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tmpt.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blkName))
                    blkId = bt[blkName];
                else
                    return;
                BlockTableRecord blk = (BlockTableRecord)tmpt.GetObject(blkId, OpenMode.ForWrite);
                ObjectIdCollection blkRefs = blk.GetBlockReferenceIds(true, true);
                if (blkRefs != null && blkRefs.Count > 0)
                {
                    foreach (ObjectId blkRefId in blkRefs)
                    {
                        BlockReference blkRef = (BlockReference)tmpt.GetObject(blkRefId, OpenMode.ForWrite);
                        blkRef.Erase();
                    }
                }
                blk.Erase();
                tmpt.Commit();
            }

        }

        static void ReversePipe(Pipe pipe)
        {
            ObjectId newStartStrust = pipe.EndStructureId;
            ObjectId newEndStrust = pipe.StartStructureId;

            pipe.Disconnect(ConnectorPositionType.Start);
            pipe.Disconnect(ConnectorPositionType.End);

            Point3d newStart = new Point3d(pipe.EndPoint.X, pipe.EndPoint.Y, pipe.EndPoint.Z);
            Point3d newEnd = new Point3d(pipe.StartPoint.X, pipe.StartPoint.Y, pipe.StartPoint.Z);

            pipe.StartPoint = newStart;
            pipe.EndPoint = newEnd;

            pipe.ConnectToStructure(ConnectorPositionType.Start, newStartStrust, false);
            pipe.ConnectToStructure(ConnectorPositionType.End, newEndStrust, false);

        }

        static void CreateSlopeBlockPart(Database db, BlockTableRecord slopeBtr, Pipe pipe, Point3d point1, Point3d point2, Point3d point3, Point3d point4, double partLen)
        {
            using (Transaction tmpt = db.TransactionManager.StartTransaction())
            {
                MText partLenText = new MText();
                MText partSlopeText = new MText();
                Line slopeLine = new Line();

                partLenText.Contents = Math.Round(partLen, 1).ToString();

                if (pipe.Slope > 0)
                {
                    slopeLine = new Line(point2, point3);

                    partSlopeText.Location = new Point3d(point4.X - 1, point4.Y - 1, 0);
                    partSlopeText.Contents = $"{Math.Round(pipe.Slope * 1000)}";
                    partSlopeText.Attachment = AttachmentPoint.TopRight;

                    partLenText.Location = new Point3d(point1.X + 1, point1.Y + 1, 0);
                    partLenText.Attachment = AttachmentPoint.BottomLeft;
                }
                else
                {
                    slopeLine = new Line(point1, point4);

                    partLenText.Location = new Point3d(point3.X - 1, point3.Y + 1, 0);
                    partLenText.Attachment = AttachmentPoint.BottomRight;

                    partSlopeText.Contents = $"{Math.Round(pipe.Slope * -1000)}";
                    partSlopeText.Location = new Point3d(point2.X + 1, point2.Y - 1, 0);
                    partSlopeText.Attachment = AttachmentPoint.TopLeft;
                }

                slopeBtr.AppendEntity(partLenText);
                tmpt.AddNewlyCreatedDBObject(partLenText, true);

                slopeBtr.AppendEntity(partSlopeText);
                tmpt.AddNewlyCreatedDBObject(partSlopeText, true);

                slopeBtr.AppendEntity(slopeLine);
                tmpt.AddNewlyCreatedDBObject(slopeLine, true);

                point1 = new Point3d(point3.X, point3.Y, 0);
                point2 = new Point3d(point4.X, point4.Y, 0);

                Line endLine = new Line(point3, point4);
                slopeBtr.AppendEntity(endLine);
                tmpt.AddNewlyCreatedDBObject(endLine, true);

                tmpt.Commit();
            }
                
        }

        static void CreateTypeBlockPart(Database db, BlockTableRecord typeBtr, Pipe pipe, double privStation, double station)
        {
            using (Transaction tmpt = db.TransactionManager.StartTransaction())
            {
                MText partSize = new MText();
                partSize.Location = new Point3d(5 + privStation + (station - privStation) / 2, 7.5 / 2, 0);
                partSize.Contents = pipe.PartSizeName;
                partSize.Attachment = AttachmentPoint.MiddleCenter;
                partSize.Width = station - privStation;
                typeBtr.AppendEntity(partSize);
                tmpt.AddNewlyCreatedDBObject(partSize, true);

                tmpt.Commit();
            }
        }

    }
}
