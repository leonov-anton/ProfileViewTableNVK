using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProfileViewTableNVK
{
    public class NVKBlockTable
    {
        // константы
        const string pipeTypeText = "К2_трубы_обозначение трубы и тип изоляции-";
        const string pipeSlopeText = "К2_трубы_обозначение уклона и длины-";
        const string aligAngleText = "К2_угол_поворота_трассы-";
        const double k1Yoffset = 47.5;
        const double k2Yoffset = 40;
        const double bYoffset = 52.5;
        const double kTypeOffset = 30;
        const double bTypeOffset = 37.5;

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
            ObjectIdCollection presPipeNets = CivilDocumentPressurePipesExtension.GetPressurePipeNetworkIds(cdoc);

            using (var t = doc.TransactionManager.StartTransaction())
            {
                ObjectIdCollection selectedProfView = new ObjectIdCollection();

                foreach (ObjectId objid in psr.Value.GetObjectIds())
                {
                    if (objid.ObjectClass.Name == "AeccDbGraphProfile")
                        selectedProfView.Add(objid);
                }

                if (selectedProfView.Count == 0)
                {
                    ed.WriteMessage("\nВиды профиля не выбраны");
                    t.Commit();
                    return;
                }

                foreach (ObjectId prof_view_id in selectedProfView)
                {
                    ProfileView profileView = (ProfileView)t.GetObject(prof_view_id, OpenMode.ForRead);
                    Alignment profileViewAligment = (Alignment)t.GetObject(profileView.AlignmentId, OpenMode.ForRead);
                    
                    // выбор имен труб с переопределенным стилем
                    List<string> overrideStylePipesName = new List<string>();
                    foreach (PipeOverride pipe in profileView.PipeOverrides)
                    {
                        if (pipe.UseOverrideStyle)
                        {
                            overrideStylePipesName.Add(pipe.PipeName);
                        }
                    }

                    // коллекция труб для построения блоков
                    DBObjectCollection pipes = new DBObjectCollection();

                    // удаление ранее созданных блоков
                    BlockErase(db, $"{pipeTypeText}{profileViewAligment.Name}");
                    BlockErase(db, $"{pipeSlopeText}{profileViewAligment.Name}");
                    BlockErase(db, $"{aligAngleText}{profileViewAligment.Name}");

                    // текущее пространство документа
                    BlockTableRecord btr = (BlockTableRecord)t.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // пространсто блоков
                    BlockTable bt = (BlockTable)t.GetObject(db.BlockTableId, OpenMode.ForWrite);

                    // создание блока строки уклонов труб
                    BlockTableRecord slopeBtr = new BlockTableRecord();
                    slopeBtr.Name = $"{pipeSlopeText}{profileViewAligment.Name}";
                    ObjectId slopeBtrId = bt.Add(slopeBtr);
                    t.AddNewlyCreatedDBObject(slopeBtr, true);

                    // создание блока строки типов труб
                    BlockTableRecord typeBtr = new BlockTableRecord();
                    typeBtr.Name = $"{pipeTypeText}{profileViewAligment.Name}";
                    ObjectId typeBtrId = bt.Add(typeBtr);
                    t.AddNewlyCreatedDBObject(typeBtr, true);

                    // создание точек начало и конца разделителей в строках
                    Point3d point1 = new Point3d(0, 0, 0);
                    Point3d point2 = new Point3d(0, 5, 0);
                    Point3d point3 = new Point3d(0, 0, 0);
                    Point3d point4 = new Point3d(0, 0, 0);

                    // длина участка труб с одинаковым уклоном
                    double partLen = 0;
                    // питек окончания участка труб общего типа
                    double privStation = 0;

                    // логика для выидов профилей напорных сетей
                    if (profileView.Name.Contains("В1") || profileView.Name.Contains("В2"))
                    {
                        // создание и втавка в блок уклонов начального разделителя
                        Line startLine = new Line(point1, point2);
                        slopeBtr.AppendEntity(startLine);
                        t.AddNewlyCreatedDBObject(startLine, true);

                        // создание начального разделителя блока типов труб
                        Line startLineType = new Line(new Point3d(0, 0, 0), new Point3d(0, 7.5, 0));
                        typeBtr.AppendEntity(startLineType);
                        t.AddNewlyCreatedDBObject(startLineType, true);

                        // создание блока углов поворота трассы
                        BlockTableRecord angleBtr = new BlockTableRecord();
                        angleBtr.Name = $"{aligAngleText}{profileViewAligment.Name}";
                        ObjectId angleBtrId = bt.Add(angleBtr);
                        t.AddNewlyCreatedDBObject(angleBtr, true);

                        // создение осевой линии
                        Line centerLine = new Line(new Point3d(0, 2.5, 0), new Point3d(profileViewAligment.EndingStation, 2.5, 0));
                        centerLine.Color = Color.FromRgb(0, 0, 255);
                        centerLine.LineWeight = LineWeight.LineWeight030;
                        angleBtr.AppendEntity(centerLine);
                        t.AddNewlyCreatedDBObject(centerLine, true);

                        // части трассы
                        AlignmentEntityCollection entCol = profileViewAligment.Entities;

                        for (int i = 1; i < entCol.Count(); i++)
                        {
                            // принимаем что все части трассы только прямые
                            AlignmentLine aligEnt = (AlignmentLine)entCol[i];
                            AlignmentLine privAligEnt = (AlignmentLine)entCol[i-1];

                            // вектор части трассы
                            MyVector2d privEntVec = new MyVector2d(privAligEnt.StartPoint.X, privAligEnt.StartPoint.Y, privAligEnt.EndPoint.X, privAligEnt.EndPoint.Y);
                            MyVector2d entVec = new MyVector2d(aligEnt.StartPoint.X, aligEnt.StartPoint.Y, aligEnt.EndPoint.X, aligEnt.EndPoint.Y);

                            // угол между частями
                            double angle = Math.Round(MyVector2d.Angle(privEntVec, entVec), 1);

                            // вертикальный отрезок на углу поворота трассы
                            Line angleLine = new Line(new Point3d(aligEnt.StartStation, 0, 0), new Point3d(aligEnt.StartStation, 5, 0));
                            angleBtr.AppendEntity(angleLine);
                            t.AddNewlyCreatedDBObject(angleLine, true);

                            // кружок на углу поворота трассы
                            Circle angleCircle = new Circle();
                            angleCircle.Center = new Point3d(aligEnt.StartStation, 2.5, 0);
                            angleCircle.Radius = 0.2;
                            angleBtr.AppendEntity(angleCircle);
                            t.AddNewlyCreatedDBObject(angleCircle, true);

                            ObjectIdCollection hatchCircCol = new ObjectIdCollection();
                            hatchCircCol.Add(angleCircle.ObjectId);
                            
                            // штриховка кружка
                            Hatch hatchCirc = new Hatch();
                            hatchCirc.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                            hatchCirc.AppendLoop(HatchLoopTypes.Outermost, hatchCircCol);
                            angleBtr.AppendEntity(hatchCirc);
                            t.AddNewlyCreatedDBObject(hatchCirc, true);

                            // номер угла поворота
                            MText angleNumText = new MText();
                            angleNumText.Contents = $"УП{i}";
                            angleNumText.Location = new Point3d(aligEnt.StartStation - 0.5, 3, 0);
                            angleNumText.Attachment = AttachmentPoint.BottomRight;
                            angleBtr.AppendEntity(angleNumText);
                            t.AddNewlyCreatedDBObject(angleNumText, true);

                            // текст угла поворота
                            MText angleDigText = new MText();
                            angleDigText.Location = new Point3d(aligEnt.StartStation - 0.5, 2, 0);
                            angleDigText.Attachment = AttachmentPoint.TopRight;
                            
                            // стрелка и дуга на углу поворота
                            Polyline arow = new Polyline();
                            Arc arc = new Arc();
                            arc.Center = angleCircle.Center;
                            arc.Radius = 0.6;

                            // положение стрелки и дуги в зависимости от угла
                            if (angle > 0)
                            {
                                angleDigText.Contents = $"{angle}°";
                                arow.AddVertexAt(0, new Point2d(aligEnt.StartStation - 0.3, 3.5), 0, 0, 0);
                                arow.AddVertexAt(1, new Point2d(aligEnt.StartStation, 4.7), 0, 0, 0);
                                arow.AddVertexAt(0, new Point2d(aligEnt.StartStation + 0.3, 3.5), 0, 0, 0);
                                arow.Closed = true;
                                arc.StartAngle = 0;
                                arc.EndAngle = 1.5708;
                            }
                            else
                            {
                                angleDigText.Contents = $"{-angle}°";
                                arow.AddVertexAt(0, new Point2d(aligEnt.StartStation - 0.3, 1.5), 0, 0, 0);
                                arow.AddVertexAt(1, new Point2d(aligEnt.StartStation, 0.3), 0, 0, 0);
                                arow.AddVertexAt(0, new Point2d(aligEnt.StartStation + 0.3, 1.5), 0, 0, 0);
                                arow.Closed = true;
                                arc.StartAngle = 4.71239;
                                arc.EndAngle = 0;
                            }

                            angleBtr.AppendEntity(angleDigText);
                            t.AddNewlyCreatedDBObject(angleDigText, true);

                            angleBtr.AppendEntity(arow);
                            t.AddNewlyCreatedDBObject(arow, true);

                            angleBtr.AppendEntity(arc);
                            t.AddNewlyCreatedDBObject(arc, true);

                            ObjectIdCollection hatchArowCol = new ObjectIdCollection();
                            hatchArowCol.Add(arow.ObjectId);

                            // штриховка стрелочки
                            Hatch hatchArow = new Hatch();
                            hatchArow.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                            hatchArow.AppendLoop(HatchLoopTypes.Outermost, hatchArowCol);
                            angleBtr.AppendEntity(hatchArow);
                            t.AddNewlyCreatedDBObject(hatchArow, true);
                        }

                        // вставка блока строки углов поворота трассы
                        BlockReference brAngle = new BlockReference(new Point3d(profileView.Location.X, profileView.Location.Y - 70, profileView.Location.Z), angleBtrId);
                        btr.AppendEntity(brAngle);
                        t.AddNewlyCreatedDBObject(brAngle, true);

                        // наполение коллекции труб для создания блоков
                        foreach (ObjectId pipeNetId in presPipeNets)
                        {
                            PressurePipeNetwork pipenet = (PressurePipeNetwork)t.GetObject(pipeNetId, OpenMode.ForRead);
                            ObjectIdCollection pipeIds = pipenet.GetPipeIds();
                            foreach (ObjectId pipeId in pipeIds)
                            {
                                PressurePipe pipeObj = (PressurePipe)t.GetObject(pipeId, OpenMode.ForWrite);
                                if (pipeObj.GetProfileViewsDisplayingMe().Contains(prof_view_id)
                                    && !pipeObj.NetworkName.ToLower().Contains("футляр")
                                    && !overrideStylePipesName.Contains(pipeObj.Name))
                                {
                                    pipes.Add(pipeObj);
                                }
                            }
                        }

                        // придыдущая труба
                        PressurePipe privPipe = null;

                        // поиск первой напорной трубы в цепочке
                        foreach (PressurePipe pipe in pipes)
                        {
                            if (Math.Round(pipe.StartPoint.X, 3) == Math.Round(profileViewAligment.StartPoint.X, 3) 
                                && Math.Round(pipe.StartPoint.Y, 3) == Math.Round(profileViewAligment.StartPoint.Y, 3))
                            {
                                pipes.Remove(pipe);
                                privPipe = pipe;
                                break;
                            }
                        }

                        // поиск последующих труб и создание блока уклонов
                        while (pipes.Count > 0)
                        {
                            ObjectId privPartId = privPipe.EndPartId;
                            foreach(PressurePipe pipe in pipes)
                            {
                                double station = Double.NaN;
                                double offset = Double.NaN;

                                profileViewAligment.StationOffset(pipe.StartPoint.X, pipe.StartPoint.Y, ref station, ref offset);

                                if (Math.Round(pipe.Slope, 3) == Math.Round(privPipe.Slope, 3) 
                                    && (Math.Round(privPipe.EndPoint.Z - privPipe.OuterDiameter / 2, 1) == Math.Round(pipe.StartPoint.Z - pipe.OuterDiameter / 2, 1) 
                                    || Math.Round(privPipe.EndPoint.Z + privPipe.OuterDiameter / 2, 1) == Math.Round(pipe.StartPoint.Z + pipe.OuterDiameter / 2, 1)))
                                {
                                    partLen += privPipe.Length2DCenterToCenter;
                                }
                                
                                else
                                {
                                    partLen += privPipe.Length2DCenterToCenter;

                                    point3 = new Point3d(station, 0, 0);
                                    point4 = new Point3d(station, 5, 0);

                                    CreateSlopeBlockPart(db, slopeBtr, privPipe.Slope, point1, point2, point3, point4, partLen);

                                    point1 = new Point3d(point3.X, point3.Y, 0);
                                    point2 = new Point3d(point4.X, point4.Y, 0);

                                    partLen = 0;
                                }

                                // создание контента блока типов труб
                                if (pipe.PartDescription != privPipe.PartDescription)
                                {
                                    CreateTypeBlockPart(db, typeBtr, null, privStation, station, privPipe);

                                    privStation = station;
                                }

                                pipes.Remove(pipe);
                                privPipe = pipe;
                                break;
                            }
                        }

                        // создание последнего разделителя строки уклонов
                        point3 = new Point3d(profileViewAligment.EndingStation, 0, 0);
                        point4 = new Point3d(profileViewAligment.EndingStation, 5, 0);

                        partLen += privPipe.Length2DCenterToCenter;

                        CreateSlopeBlockPart(db, slopeBtr, privPipe.Slope, point1, point2, point3, point4, partLen);

                        CreateTypeBlockPart(db, typeBtr, null, privStation, profileViewAligment.EndingStation, privPipe);

                    }

                    // логика для безнапортных сетей
                    else if(profileView.Name.Contains("К1") || profileView.Name.Contains("К2"))
                    {
                        // создание точек первого разделителя (таблица безнаторных сетей смещена на 5мм по сравнению с напорными)
                        point1 = new Point3d(5, 0, 0);
                        point2 = new Point3d(5, 5, 0);

                        // создание и втавка в блок уклонов начального разделителя
                        Line startLine = new Line(point1, point2);
                        slopeBtr.AppendEntity(startLine);
                        t.AddNewlyCreatedDBObject(startLine, true);

                        // создание начального разделителя блока типов труб
                        Line startLineType = new Line(new Point3d(5, 0, 0), new Point3d(5, 7.5, 0));
                        typeBtr.AppendEntity(startLineType);
                        t.AddNewlyCreatedDBObject(startLineType, true);

                        // наполение коллекции труб
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

                        // предыдущая труба
                        Pipe privPipe = null;

                        // поиск первой трубы (разворот при необходимости)
                        foreach (Pipe pipe in pipes)
                        {
                            if (pipe.StartPoint.X == profileViewAligment.StartPoint.X && pipe.StartPoint.Y == profileViewAligment.StartPoint.Y)
                            {
                                pipes.Remove(pipe);
                                privPipe = pipe;
                                break;
                            } 
                            else if (pipe.EndPoint.X == profileViewAligment.StartPoint.X && pipe.EndPoint.Y == profileViewAligment.StartPoint.Y)
                            {
                                ReversePipe(pipe);
                                pipes.Remove(pipe);
                                privPipe = pipe;
                                break;
                            }
                        }

                        // обработка последующих труб
                        while (pipes.Count > 0)
                        {
                            ObjectId privStructId = privPipe.EndStructureId;
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

                                    if (Math.Round(pipe.Slope, 3) == Math.Round(privPipe.Slope, 3) 
                                        && (Math.Round(privPipe.EndPoint.Z - privPipe.InnerDiameterOrWidth / 2, 1) == Math.Round(pipe.StartPoint.Z - pipe.InnerDiameterOrWidth / 2, 1) 
                                        || Math.Round(privPipe.EndPoint.Z + privPipe.InnerDiameterOrWidth / 2, 1) == Math.Round(pipe.StartPoint.Z + pipe.InnerDiameterOrWidth / 2, 1)))
                                    {
                                        partLen += privPipe.Length2D;
                                    }
                                    else
                                    {
                                        partLen += privPipe.Length2D;
                                        
                                        point3 = new Point3d(5 + station, 0, 0);
                                        point4 = new Point3d(5 + station, 5, 0);

                                        CreateSlopeBlockPart(db, slopeBtr, privPipe.Slope, point1, point2, point3, point4, partLen);

                                        point1 = new Point3d(point3.X, point3.Y, 0);
                                        point2 = new Point3d(point4.X, point4.Y, 0);

                                        partLen = 0;

                                    }

                                    if (pipe.PartSizeName != privPipe.PartSizeName)
                                    {
                                        CreateTypeBlockPart(db, typeBtr, privPipe, privStation, station, null);

                                        privStation = station;
                                    }

                                    pipes.Remove(pipe);
                                    privPipe = pipe;
                                    break;

                                }
                            }
                        }

                        point3 = new Point3d(5 + profileViewAligment.EndingStation, 0, 0);
                        point4 = new Point3d(5 + profileViewAligment.EndingStation, 5, 0);

                        partLen += privPipe.Length2D;

                        CreateSlopeBlockPart(db, slopeBtr, privPipe.Slope, point1, point2, point3, point4, partLen);

                        CreateTypeBlockPart(db, typeBtr, privPipe, privStation, profileViewAligment.EndingStation, null);

                    }
                    else
                    {
                        ed.WriteMessage("\nНе выбраны виды профиля водопровода или канализации");
                        t.Commit();
                        return;
                    }

                    // определение смещения точки вставки по Y
                    double slopeYOffset;
                    double typeYOffset = kTypeOffset;
                    if (profileView.Name.Contains("В1") || profileView.Name.Contains("В2"))
                    {
                        slopeYOffset = bYoffset;
                        typeYOffset = bTypeOffset;
                    }       
                    else if (profileView.Name.Contains("К1"))
                        slopeYOffset = k1Yoffset;
                    else
                        slopeYOffset = k2Yoffset;

                    // вставка блока строки уклона труб
                    BlockReference brSlope = new BlockReference(new Point3d(profileView.Location.X, profileView.Location.Y - slopeYOffset, profileView.Location.Z), slopeBtrId);
                    btr.AppendEntity(brSlope);
                    t.AddNewlyCreatedDBObject(brSlope, true);

                    // вставка блока строки типа труб
                    BlockReference brType = new BlockReference(new Point3d(profileView.Location.X, profileView.Location.Y - typeYOffset, profileView.Location.Z), typeBtrId);
                    btr.AppendEntity(brType);
                    t.AddNewlyCreatedDBObject(brType, true);
                }
                t.Commit();
            }

        }

        static void BlockErase(Database db, string blkName)
        // удаление блока и всех вхождений блока
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
                    // удаление вхождений блока
                    foreach (ObjectId blkRefId in blkRefs)
                    {
                        BlockReference blkRef = (BlockReference)tmpt.GetObject(blkRefId, OpenMode.ForWrite);
                        blkRef.Erase();
                    }
                }
                // удаление объекта блока
                blk.Erase();
                tmpt.Commit();
            }

        }

        static void ReversePipe(Pipe pipe)
        // разворот безнапорной трубы
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

        static void CreateSlopeBlockPart(Database db, BlockTableRecord slopeBtr, double pipeSlope, Point3d point1, Point3d point2, Point3d point3, Point3d point4, double partLen)
        // создание блока строки уклона безнапорной трубы
        {

            using (Transaction tmpt = db.TransactionManager.StartTransaction())
            {
                MText partLenText = new MText();
                MText partSlopeText = new MText();
                Line slopeLine = new Line();

                partLenText.Contents = Math.Round(partLen, 1).ToString();

                if (pipeSlope > 0)
                {
                    slopeLine = new Line(point2, point3);

                    partSlopeText.Location = new Point3d(point4.X - 1, point4.Y - 1, 0);
                    partSlopeText.Contents = $"{Math.Round(pipeSlope * 1000)}";
                    partSlopeText.Attachment = AttachmentPoint.TopRight;

                    partLenText.Location = new Point3d(point1.X + 1, point1.Y + 1, 0);
                    partLenText.Attachment = AttachmentPoint.BottomLeft;
                }
                else
                {
                    slopeLine = new Line(point1, point4);

                    partLenText.Location = new Point3d(point3.X - 1, point3.Y + 1, 0);
                    partLenText.Attachment = AttachmentPoint.BottomRight;

                    partSlopeText.Contents = $"{Math.Round(pipeSlope * -1000)}";
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

        static void CreateTypeBlockPart(Database db, BlockTableRecord typeBtr, Pipe pipe, double privStation, double station, PressurePipe presPipe)

        {
            using (Transaction tmpt = db.TransactionManager.StartTransaction())
            {
                Line nextLine = null;

                MText partSize = new MText();
                partSize.Location = new Point3d(5 + privStation + (station - privStation) / 2, 7.5 / 2, 0);
                if(pipe != null)
                { 
                    partSize.Contents = pipe.PartSizeName;
                    nextLine = new Line(new Point3d(5 + station, 0, 0), new Point3d(5 + station, 7.5, 0));
                }
                else
                {
                    partSize.Contents = presPipe.PartDescription;
                    nextLine = new Line(new Point3d(station, 0, 0), new Point3d(station, 7.5, 0));
                } 
                partSize.Attachment = AttachmentPoint.MiddleCenter;
                partSize.Width = station - privStation;
                typeBtr.AppendEntity(partSize);
                tmpt.AddNewlyCreatedDBObject(partSize, true);

                typeBtr.AppendEntity(nextLine);
                tmpt.AddNewlyCreatedDBObject(nextLine, true);

                tmpt.Commit();
            }
        }
    }
}
