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
        public bool WrapTrackingChanges = false;
        public WmlComparerWrapTrackingChangesSettings WrapTrackingChangesSettings = new WmlComparerWrapTrackingChangesSettings();
    }

    public class WmlComparerWrapTrackingChangesSettings
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

        class ComparisonInitAtomsGroupInfo
        {
            public IList<ComparisonUnitAtom> Atoms;
            public string Status;
            public int Position;
            public int Size;
        
            public bool IsChanged { 
                get => Status == "Deleted" || Status == "Inserted"; 
            }
        }

        private static void MarkInsertedDeletedComparisonUnitAtomsBounds(
            List<ComparisonUnitAtom> comparisonUnitAtoms,
            WmlComparerInternalSettings internalSettings   
        ) 
        {
            int getPrevChangedAtomsGroupPositionDistance(List<ComparisonInitAtomsGroupInfo> atomsGroups, int index)
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

            int getNextChangedAtomsGroupPositionDistance(List<ComparisonInitAtomsGroupInfo> atomsGroups, int index)
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
                    var groups = GroupAdjacentComparisonUnitAtomsByTrackedChange(comparisonUnitAtoms, -1)
                        .Select(group => {
                            var spl = group.Key.Split('|');
                            var status = spl[1];
                            var g = new ComparisonInitAtomsGroupInfo() 
                            {
                                Atoms = group.ToList(),
                                Status = status, 
                                Position = position,
                                Size = group.Count(), 
                            };
                            position += g.Size;
                            return g;
                        })
                        .ToList();
                    
                    
                    for (var i = 0; i < groups.Count; i++) 
                    {
                        var group = groups[i];

                        if (group.IsChanged)
                        {
                            var changeGroupUnid = Util.GenerateUnid();

                            var firstAtom = group.Atoms.FirstOrDefault(a => a.ContentElement.Name == W.t);
                            if (firstAtom != null)
                            {
                                firstAtom.ChangeGroupStart = true;
                                firstAtom.ChangeGroupUnid = changeGroupUnid;

                                var prevChangedGroupDistance = getPrevChangedAtomsGroupPositionDistance(groups, i);
                                if (prevChangedGroupDistance != -1 && 
                                    prevChangedGroupDistance < internalSettings.MergeSettings.WrapTrackingChangesSettings.FragmentDistance
                                )
                                {
                                    firstAtom.ChangeGroupRequiresHighlight = true;
                                }
                            }

                            var lastAtom = group.Atoms.LastOrDefault(a => a.ContentElement.Name == W.t);
                            if (lastAtom != null)
                            {
                                lastAtom.ChangeGroupEnd = true;
                                lastAtom.ChangeGroupUnid = changeGroupUnid;

                                var nextChangedGroupDistance = getNextChangedAtomsGroupPositionDistance(groups, i);
                                if (nextChangedGroupDistance != -1 && 
                                    nextChangedGroupDistance < internalSettings.MergeSettings.WrapTrackingChangesSettings.FragmentDistance
                                )
                                {
                                    lastAtom.ChangeGroupRequiresHighlight = true;
                                }
                            }
                        }
                    };
            });
        }
    }
}