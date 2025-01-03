using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class WmlComparerMergeSettings
    {
        public bool FormatTrackingChanges = false;
        public WmlComparerFormatTrackingChangesSettings FormatTrackingChangesSettings = new WmlComparerFormatTrackingChangesSettings();
    }

    public class WmlComparerFormatTrackingChangesSettings
    {
        public string FragmentStart = "[";
        public string FragmentEnd = "]";
        public int FragmentDistance = 10;
    }

    public static partial class WmlComparer
    {
        public static WmlDocument Merge(WmlDocument original,
            IEnumerable<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings, WmlComparerMergeSettings mergeSettings)
        {
            var originalWithUnids = PreProcessMarkup(original);
            WmlDocument merged = new WmlDocument(originalWithUnids);

            var revisedDocumentInfoListCount = revisedDocumentInfoList.Count();
            var revisors = revisedDocumentInfoList.Select(r => r.Revisor).ToArray();

            using (MemoryStream mergedMs = new MemoryStream())
            {
                mergedMs.Write(merged.DocumentByteArray, 0, merged.DocumentByteArray.Length);
                using (WordprocessingDocument mergedWDoc = WordprocessingDocument.Open(mergedMs, true))
                {
                    var mergedMainDocPart = mergedWDoc.MainDocumentPart;
                    var mergedMainDocPartXDoc = mergedMainDocPart.GetXDocument();

                    var consolidatedByUnid = mergedMainDocPartXDoc
                        .Descendants()
                        .Where(d => (d.Name == W.p || d.Name == W.tbl) && d.Attribute(PtOpenXml.Unid) != null)
                        .ToDictionary(d => (string)d.Attribute(PtOpenXml.Unid));

                    var revisedDocumentInfoList2 = revisedDocumentInfoList.ToList();

                    for (var i = 0; i < revisedDocumentInfoList2.Count; i++)
                    {
                        var revisedDocumentInfo = revisedDocumentInfoList2[i];
                        var isLast = i == revisedDocumentInfoList2.Count - 1;

                        var internalSettings = new WmlComparerInternalSettings(settings, mergeSettings)
                        {
                            PreProcessMarkupInOriginal = false,
                            RevisionsAmount = revisedDocumentInfoList2.Count,
                            MergeMode = true, 
                            MergeIteration = i,
                            // for last item, we need to resolve all accumulated tracking changes
                            ResolveTrackingChanges = isLast,
                            MergeRevisors = revisors,
                        };

                        var revised = revisedDocumentInfo.RevisedDocument;
                        var delta = WmlComparer.CompareInternal(merged, revised, internalSettings);

                        var colorRgb = revisedDocumentInfo.Color.ToArgb();
                        var colorString = colorRgb.ToString("X");
                        if (colorString.Length == 8)
                            colorString = colorString.Substring(2);

                        using (MemoryStream msOriginalWithUnids = new MemoryStream())
                        using (MemoryStream msDelta = new MemoryStream())
                        {
                            msOriginalWithUnids.Write(originalWithUnids.DocumentByteArray, 0, originalWithUnids.DocumentByteArray.Length);
                            msDelta.Write(delta.DocumentByteArray, 0, delta.DocumentByteArray.Length);
                            
                            using (WordprocessingDocument wDocOriginalWithUnids = WordprocessingDocument.Open(msOriginalWithUnids, true))
                            using (WordprocessingDocument wDocDelta = WordprocessingDocument.Open(msDelta, true))
                            {
                                var deltaMainDocPart = wDocDelta.MainDocumentPart;
                                var deltaMainDocPartXDoc = deltaMainDocPart.GetXDocument();


                                deltaMainDocPart.PutXDocument();
                            }
                        }

                        merged = delta;
                    }
                }
            }
            return merged;
        }

        private static void StoreChangeTrackingStatusForMerge(XDocument doc, int mergeIteration)
        {
            doc.Root
                .Descendants()
                .Where(e => e.Attribute(PtOpenXml.Status) != null)
                .ToList()
                .ForEach(e => {
                    var status = (string) e.Attribute(PtOpenXml.Status);
                    var mergeStatus = (string) e.Attribute(PtOpenXml.MergeStatus);

                    if (mergeStatus == null || status == "Deleted")
                    {
                        e.SetAttributeValue(PtOpenXml.MergeStatus, status);
                        
                        var assignedMergeIterations = (string) e.Attribute(PtOpenXml.MergeIterations);
                        if (assignedMergeIterations != null)
                        {
                            assignedMergeIterations += "," + mergeIteration;
                        } 
                        else
                        {
                            assignedMergeIterations = mergeIteration.ToString();
                        }
                        e.SetAttributeValue(PtOpenXml.MergeIterations, assignedMergeIterations);
                    }
                });
        }

        private static void RestoreChangeTrackingStatusForMerge(XDocument doc)
        {
            doc.Root
                .Descendants()
                .Where(e => e.Attribute(PtOpenXml.MergeStatus) != null)
                .ToList()
                .ForEach(e => {
                    var status = (string) e.Attribute(PtOpenXml.MergeStatus);
                    e.SetAttributeValue(PtOpenXml.Status, status);
                });
        }

        class ComparisonUnitAtomsGroupInfo
        {
            public IList<ComparisonUnitAtom> Atoms;
            public string Status;
            public int Position;
            public int Size;
        
            public bool IsChanged { 
                get => Status == "Deleted" || Status == "Inserted"; 
            }
        }

        private static void MarkInsertedDeletedComparisonUnitAtomsBoundsForFormatting(
            List<ComparisonUnitAtom> comparisonUnitAtoms,
            WmlComparerInternalSettings internalSettings   
        ) 
        {
            int getPrevChangedAtomsGroupPositionDistance(List<ComparisonUnitAtomsGroupInfo> atomsGroups, int index)
            {
                var currentAtomGroup = atomsGroups[index];
                for (var i = index - 1; i >= 0; i--)
                {
                    var atomGroup = atomsGroups[i];
                    if (atomGroup.IsChanged)
                    {
                        return currentAtomGroup.Position - (atomGroup.Position + atomGroup.Size);
                    }
                }
                return -1;
                
            }

            int getNextChangedAtomsGroupPositionDistance(List<ComparisonUnitAtomsGroupInfo> atomsGroups, int index)
            {
                var currentAtomGroup = atomsGroups[index];
                for (var i = index + 1; i < atomsGroups.Count; i++)
                {
                    var atomGroup = atomsGroups[i];
                    if (atomGroup.IsChanged)
                    {
                        return atomGroup.Position - (currentAtomGroup.Position + currentAtomGroup.Size);
                    }
                }
                return -1;
            }

            comparisonUnitAtoms
                // group atoms by paragraph
                .GroupAdjacent(cua => 
                    cua.AncestorElements
                        .Select((ae, i) => new { Element = ae, Index = i })
                        .Where(aei => aei.Element.Name == W.p)
                        .Select(aei => cua.AncestorUnids[aei.Index])
                        .FirstOrDefault()
                    )
                 .Where(g => g.Key != null)
                 .ToList()
                 .ForEach(paraGroup => {
                    var position = 0;

                    // group texts within each paragraphs by tracking changes
                    var changedGroups = GroupAdjacentComparisonUnitAtomsByTrackedChange(paraGroup, -1, true)
                        .Select(group => {
                            var firstGroupAtom = group.FirstOrDefault();
                            var g = new ComparisonUnitAtomsGroupInfo() 
                            {
                                Atoms = group.ToList(),
                                // for last merged document, its Status should come from the CorrelationStatus field
                                Status = firstGroupAtom.MergeStatus ?? firstGroupAtom.CorrelationStatus.ToString(), 
                                Position = position,
                                Size = group.Count(), 
                            };
                            position += g.Size;
                            return g;
                        })
                        .Where(g => g.IsChanged)
                        .ToList();
                        
                    for (var i = 0; i < changedGroups.Count; i++) 
                    {
                        var changedGroup = changedGroups[i];
                        var firstAtom = changedGroup.Atoms.FirstOrDefault(a => a.ContentElement.Name == W.t);
                        var lastAtom = changedGroup.Atoms.LastOrDefault(a => a.ContentElement.Name == W.t);
                        var changeGroupUnid = Util.GenerateUnid();

                        if (firstAtom != null && lastAtom != null)
                        {
                            firstAtom.ChangeGroupStart = true;
                            firstAtom.ChangeGroupUnid = changeGroupUnid;

                            lastAtom.ChangeGroupEnd = true;
                            lastAtom.ChangeGroupUnid = changeGroupUnid;

                            var prevChangedGroupDistance = getPrevChangedAtomsGroupPositionDistance(changedGroups, i);
                            var nextChangedGroupDistance = getNextChangedAtomsGroupPositionDistance(changedGroups, i);

                            //Console.WriteLine(prevChangedGroupDistance);
                            if ((
                                    prevChangedGroupDistance != -1 && 
                                    prevChangedGroupDistance < internalSettings.MergeSettings.FormatTrackingChangesSettings.FragmentDistance
                                ) ||
                                (
                                    nextChangedGroupDistance != -1 && 
                                    nextChangedGroupDistance < internalSettings.MergeSettings.FormatTrackingChangesSettings.FragmentDistance
                                )
                            )
                            {
                                firstAtom.ChangeGroupRequireFormatting = true;
                                lastAtom.ChangeGroupRequireFormatting = true;
                            }
                        }
                    };
            });
        }

    }
}