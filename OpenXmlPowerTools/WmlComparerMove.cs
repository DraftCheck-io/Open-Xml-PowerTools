using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        
        private class CorrelatedSequenceStats
        {
            public readonly int Total = 0;
            public readonly int Equal = 0;
            public readonly float Percentage = 0;

            public CorrelatedSequenceStats(List<CorrelatedSequence> correlatedSequence = null)
            {
                var total1 = 0;
                var total2 = 0;
                var equal = 0;

                foreach (var cs in correlatedSequence)
                {
                    var cus = cs.ComparisonUnitArray1 ?? cs.ComparisonUnitArray2;
                    var length = cus.Sum(cu => cu.DescendantContentAtomsCount);

                    if (cs.CorrelationStatus == CorrelationStatus.Equal)
                    {
                        equal += length;
                        total1 += length;
                        total2 += length;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Deleted)
                    {
                        total1 += length;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Inserted)
                    {
                        total2 += length;
                    }
                }

                var total = Math.Min(total1, total2);
                var percentage = total > 0 ? (float)equal / total : 0;

                Total = total;
                Equal = equal;
                Percentage = percentage;
            }

        }

        private static readonly int s_MinMovedSequenceLength = 50;
        private static readonly float s_MinMovedSequenceEquityRatio = 0.5F;

        private static void AssignIndexesToAllRunElements(XElement contentParent)
        {
            var content = contentParent.Descendants(W.r);
            var rCount = 0;
            foreach (var d in content)
            {
                if (d.Attribute(PtOpenXml.Index) == null) {
                    var newAttr = new XAttribute(PtOpenXml.Index, rCount.ToString());
                    d.Add(newAttr);
                }
                rCount++;
            }
        }

        private static void AddRunIndexesToMarkupInContentParts(WordprocessingDocument wDoc)
        {
            var mdp = wDoc.MainDocumentPart.GetXDocument();
            AssignIndexesToAllRunElements(mdp.Root);
            IgnorePt14Namespace(mdp.Root);
            wDoc.MainDocumentPart.PutXDocument();

            if (wDoc.MainDocumentPart.FootnotesPart != null)
            {
                var p = wDoc.MainDocumentPart.FootnotesPart.GetXDocument();
                AssignIndexesToAllRunElements(p.Root);
                IgnorePt14Namespace(p.Root);
                wDoc.MainDocumentPart.FootnotesPart.PutXDocument();
            }

            if (wDoc.MainDocumentPart.EndnotesPart != null)
            {
                var p = wDoc.MainDocumentPart.EndnotesPart.GetXDocument();
                AssignIndexesToAllRunElements(p.Root);
                IgnorePt14Namespace(p.Root);
                wDoc.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static void DetectMovedContentInCorrelatedSequence(IEnumerable<CorrelatedSequence> correlatedSequence, WmlComparerInternalSettings internalSettings)
        {
            void backupComparisonUnitAtomsUnids(IEnumerable<ComparisonUnitAtom> comparisonUnitAtoms) 
            {
                void doBackupAttribute(XElement element, XName attributeName) 
                {
                    XName backupAttributeName = attributeName + "Backup";
                    var attr = element.Attribute(attributeName);
                    var backupAttr = element.Attribute(backupAttributeName);
                    if (attr != null && backupAttr == null)
                    {
                        element.Add(new XAttribute(backupAttributeName, attr.Value));
                    }
                }

                foreach (var cua in comparisonUnitAtoms)
                {
                    doBackupAttribute(cua.ContentElement, PtOpenXml.Unid);

                    foreach (var ae in cua.AncestorElements)
                    {
                        doBackupAttribute(ae, PtOpenXml.Unid);
                    }

                    if (cua.ContentElement.Attribute(PtOpenXml.Unid2) == null)
                    {
                        cua.ContentElement.Add(new XAttribute(PtOpenXml.Unid2, Util.GenerateUnid()));
                    }
                }
            }

            void restoreComparisonUnitAtomsUnids(IEnumerable<ComparisonUnitAtom> comparisonUnitAtoms) 
            {
                void doRestoreAttribute(XElement element, XName attributeName) 
                {
                    XName backupAttributeName = attributeName + "Backup";
                    var attr = element.Attribute(attributeName);
                    var backupAttr = element.Attribute(backupAttributeName);
                    if (attr != null && backupAttr != null)
                    {
                        attr.Value = backupAttr.Value;
                        backupAttr.Remove();
                    }
                }

                foreach (var cua in comparisonUnitAtoms)
                {
                    doRestoreAttribute(cua.ContentElement, PtOpenXml.Unid);

                    foreach (var ae in cua.AncestorElements)
                    {
                        doRestoreAttribute(ae, PtOpenXml.Unid);
                    }

                    cua.ContentElement.Attribute(PtOpenXml.Unid2)?.Remove();
                }
            }

            void assignMoveFromUnidToComparisonUnitAtoms(
                IEnumerable<ComparisonUnit> comparisonUnits, 
                string moveUid, 
                CorrelationStatus moveStatus,
                int moveFragmentIndex
            ) 
            {
                foreach (var cu in comparisonUnits)
                    foreach (var ca in cu.DescendantContentAtoms())
                    {
                        ca.MoveFromUnid = moveUid;
                        ca.MoveStatus = moveStatus;
                        ca.MoveFragmentIndex = moveFragmentIndex;
                    }
            }

            void assignMoveToUnidToComparisonUnitAtoms(
                IEnumerable<ComparisonUnit> comparisonUnits, 
                string moveUid,
                CorrelationStatus moveStatus,
                int moveFragmentIndex
            ) 
            {
                foreach (var cu in comparisonUnits)
                    foreach (var ca in cu.DescendantContentAtoms())
                    {
                        ca.MoveToUnid = moveUid;
                        ca.MoveStatus = moveStatus;
                        ca.MoveFragmentIndex = moveFragmentIndex;
                    }
            }

            IEnumerable<string> collectComparisonUnitAtomsUnids(IEnumerable<ComparisonUnit> comparisonUnits)
            {
                return comparisonUnits
                    .SelectMany(cu => cu
                        .DescendantContentAtoms()
                        .Select(a => a.ContentElement.Attribute(PtOpenXml.Unid2)?.Value)
                    );
            }

            IEnumerable<ComparisonUnit> reassembleComparisonUnits(
                IEnumerable<ComparisonUnit> comparisonUnits,
                IEnumerable<string> excludedAtomsUnids = null,
                bool excludeNonTextAtomElements = true
            )
            {
                var result = new List<ComparisonUnit>();

                foreach (var cu in comparisonUnits)
                {
                    if (cu is ComparisonUnitAtom cua)
                    {
                        // skip if content element is not a text element
                        if (excludeNonTextAtomElements && cua.ContentElement.Name != W.t && cua.ContentElement.Name != W.delText)
                            continue;

                        if (excludedAtomsUnids != null)
                        {
                            var unid = cua.ContentElement.Attribute(PtOpenXml.Unid2)?.Value;
                            if (unid == null || !excludedAtomsUnids.Contains(unid))
                                result.Add(cu);
                        }
                        else
                        {
                            result.Add(cu);
                        }
                    }
                    else
                    {
                        var contents = reassembleComparisonUnits(cu.Contents, excludedAtomsUnids, excludeNonTextAtomElements);
                        if (contents.Any())
                        {
                            if (cu is ComparisonUnitWord cuw)
                                result.Add(new ComparisonUnitWord(contents.OfType<ComparisonUnitAtom>()));
                            // Do not process ComparisonUnitGroup here (they were unwrapped before during simplification)
                            // else if (cu is ComparisonUnitGroup cug)
                            //     result.Add(new ComparisonUnitGroup(contents, cug.ComparisonUnitGroupType, cug.Level));
                            else
                                // do not expect ComparisonUnitGroup here (they were unwrapped before)
                                throw new OpenXmlPowerToolsException("Internal error: unexpected ComparisonUnit type");
                        }
                    }
                }

                return result;
            }

            IEnumerable<ComparisonUnit> unwrapComparisonUnitGroups(IEnumerable<ComparisonUnit> comparisonUnits)
            {
                return comparisonUnits.SelectMany(
                    cu => (cu is ComparisonUnitGroup cug)
                        ? unwrapComparisonUnitGroups(cug.Contents)
                        : new ComparisonUnit[] { cu }
                );
            }

            IEnumerable<IEnumerable<ComparisonUnit>> getComparisonUnitsChunksByStatus(
                IEnumerable<CorrelatedSequence> correlatedSequence2,
                CorrelationStatus status,
                int minLength = 0
            )
            {
                return correlatedSequence2
                    .Where(cs => cs.CorrelationStatus == status)
                    .Select(cs => cs.ComparisonUnitArray1 ?? cs.ComparisonUnitArray2)
                    // unwrap ComparisonUnitGroups for simplification
                    .Select(cua => unwrapComparisonUnitGroups(cua))
                    .Select(cus => reassembleComparisonUnits(cus))
                    // additionally split comparison units into chunks by paragraphs
                    .SelectMany(cus => cus
                        .GroupAdjacent(cu => cu
                            .DescendantContentAtoms()
                            .First()
                            .AncestorElements
                            .FirstOrDefault(ae => ae.Name == W.p)
                            .Attribute(PtOpenXml.Unid)?.Value
                        )
                    )
                    .Where(cu => cu.Sum(c => c.DescendantContentAtomsCount) > minLength);
            }


            var deletedComparisonUnitsChunks = getComparisonUnitsChunksByStatus(correlatedSequence, CorrelationStatus.Deleted, s_MinMovedSequenceLength);
            var insertedComparisonUnitsChunks = getComparisonUnitsChunksByStatus(correlatedSequence, CorrelationStatus.Inserted, s_MinMovedSequenceLength);

            var deletedComparisonUnitAtoms = deletedComparisonUnitsChunks.SelectMany(cus => cus.SelectMany(cu => cu.DescendantContentAtoms()));
            var insertedComparisonUnitAtoms = insertedComparisonUnitsChunks.SelectMany(cus => cus.SelectMany(cu => cu.DescendantContentAtoms()));

            // Lcs algorithm changes the original Unids of the elements; so need to backup and restore them later
            backupComparisonUnitAtomsUnids(insertedComparisonUnitAtoms);
            backupComparisonUnitAtomsUnids(deletedComparisonUnitAtoms);

            var movedSequencesListWithStats = deletedComparisonUnitsChunks
                .SelectMany(deletedComparisonUnits => insertedComparisonUnitsChunks
                    .Select(insertedComparisonUnits => {
                        var lcs = Lcs(deletedComparisonUnits.ToArray(), insertedComparisonUnits.ToArray(), internalSettings);
                        return new
                        {
                            Sequences = lcs,
                            Stats = new CorrelatedSequenceStats(lcs),
                            DeletedComparisonUnits = deletedComparisonUnits,
                            InsertedComparisonUnits = insertedComparisonUnits,
                        };
                    })
                );

            while (true)
            {
                // recalculate stats and re-filter sequences
                movedSequencesListWithStats = movedSequencesListWithStats
                    .Where(ms => ms.Stats.Total > s_MinMovedSequenceLength && ms.Stats.Percentage > s_MinMovedSequenceEquityRatio)
                    .ToList()
                    .OrderByDescending(ms => ms.Stats.Equal)
                    .ThenByDescending(ms => ms.Stats.Percentage);

                var longestMovedSequencesWithStats = movedSequencesListWithStats.FirstOrDefault();

                if (longestMovedSequencesWithStats == null)
                    break;

                var moveUnid = Util.GenerateUnid();

                var movedSequences = longestMovedSequencesWithStats.Sequences;

                // strip non-equal sequences at the beginning and at the end
                var firstNonEqualIndex = 0;
                var lastNonEqualIndex = movedSequences.Count - 1;

                while (firstNonEqualIndex < lastNonEqualIndex && movedSequences[firstNonEqualIndex].CorrelationStatus != CorrelationStatus.Equal)
                    firstNonEqualIndex++;
                while (lastNonEqualIndex > firstNonEqualIndex && movedSequences[lastNonEqualIndex].CorrelationStatus != CorrelationStatus.Equal)
                    lastNonEqualIndex--;

                movedSequences = movedSequences
                    .Skip(firstNonEqualIndex)
                    .Take(lastNonEqualIndex - firstNonEqualIndex + 1)
                    .ToList();

                var movedFromComparisonUnitAtomsUnids = new List<string>();
                var movedToComparisonUnitAtomsUnids = new List<string>();

                for (int i = 0; i < movedSequences.Count; i++)
                {
                    var cs = movedSequences[i];

                    if (cs.CorrelationStatus == CorrelationStatus.Equal)
                    {
                        assignMoveFromUnidToComparisonUnitAtoms(cs.ComparisonUnitArray1, moveUnid, CorrelationStatus.Equal, i);
                        assignMoveToUnidToComparisonUnitAtoms(cs.ComparisonUnitArray2, moveUnid, CorrelationStatus.Equal, i);
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Deleted)
                    {
                        assignMoveFromUnidToComparisonUnitAtoms(cs.ComparisonUnitArray1, moveUnid, CorrelationStatus.Deleted, i);
                    } else if (cs.CorrelationStatus == CorrelationStatus.Inserted)
                    {
                        assignMoveToUnidToComparisonUnitAtoms(cs.ComparisonUnitArray2, moveUnid, CorrelationStatus.Inserted, i);
                    }

                    // consider all atoms in the moved sequences as moved (including inserted and deleted atoms inside the moved sequences)
                    if (cs.ComparisonUnitArray1 != null)
                        movedFromComparisonUnitAtomsUnids.AddRange(collectComparisonUnitAtomsUnids(cs.ComparisonUnitArray1));
                    if (cs.ComparisonUnitArray2 != null)
                        movedToComparisonUnitAtomsUnids.AddRange(collectComparisonUnitAtomsUnids(cs.ComparisonUnitArray2));
                }

                // remove moved atoms from other sequences and re-calculate them if necessary
                movedSequencesListWithStats = movedSequencesListWithStats
                    .Select(mcs => {
                        var deletedComparisonUnitAtomsUnids = collectComparisonUnitAtomsUnids(mcs.DeletedComparisonUnits);
                        var insertedComparisonUnitAtomsUnids = collectComparisonUnitAtomsUnids(mcs.InsertedComparisonUnits);

                        if (deletedComparisonUnitAtomsUnids.Overlaps(movedFromComparisonUnitAtomsUnids) ||
                            insertedComparisonUnitAtomsUnids.Overlaps(movedToComparisonUnitAtomsUnids))
                        {
                            var newDeletedComparisonUnits = reassembleComparisonUnits(mcs.DeletedComparisonUnits, movedFromComparisonUnitAtomsUnids);
                            var newInsertedComparisonUnits = reassembleComparisonUnits(mcs.InsertedComparisonUnits, movedToComparisonUnitAtomsUnids);

                            var newDeletedComparisonUnitsCount = newDeletedComparisonUnits.Sum(cu => cu.DescendantContentAtomsCount);
                            var newInsertedComparisonUnitsCount = newInsertedComparisonUnits.Sum(cu => cu.DescendantContentAtomsCount);

                            if (newDeletedComparisonUnitsCount < s_MinMovedSequenceLength || newInsertedComparisonUnitsCount < s_MinMovedSequenceLength)
                                return null;

                            var newLcs = Lcs(newDeletedComparisonUnits.ToArray(), newInsertedComparisonUnits.ToArray(), internalSettings);

                            return new
                            {
                                Sequences = newLcs,
                                Stats = new CorrelatedSequenceStats(newLcs),
                                DeletedComparisonUnits = newDeletedComparisonUnits,
                                InsertedComparisonUnits = newInsertedComparisonUnits,
                            };
                        }

                        return mcs;
                    })
                    .Where(mcs => mcs != null);
            }

            restoreComparisonUnitAtomsUnids(insertedComparisonUnitAtoms);
            restoreComparisonUnitAtomsUnids(deletedComparisonUnitAtoms);
        }

        private static void MarkMoveFromAndMoveToRanges(XNode node, WmlComparerInternalSettings internalSettings)
        {
            var settings = internalSettings.ComparerSettings;

            void markMoveRanges(XElement c, XName moveName, XName moveGroupAttrName, XName moveRangeStartName, XName moveRangeEndName = null)
            {
                var groupedMoves = c
                    .Descendants()
                    .Where(e => e.Name == moveName && !e.Ancestors(W.pPr).Any())
                    .GroupBy(e => (string)e.Attribute(moveGroupAttrName));

                foreach (var gm in groupedMoves)
                {
                    var name = gm.Key.Substring(0, 8);
                    var moves = gm.ToList();
                    var id = s_MaxId++;
                    moves
                        .First()
                        .AddBeforeSelf(new XElement(
                            moveRangeStartName,
                            new XAttribute(W.id, id),
                            new XAttribute(W.name, name),
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.date, settings.DateTimeForRevisions)
                        ));
                    moves
                        .Last()
                        .AddAfterSelf(new XElement(
                            moveRangeEndName,
                            new XAttribute(W.id, id)
                        ));
                }
            }

            XElement container = node as XElement;
            markMoveRanges(container, W.moveFrom, PtOpenXml.MoveFromUnid, W.moveFromRangeStart, W.moveFromRangeEnd);
            markMoveRanges(container, W.moveTo, PtOpenXml.MoveToUnid, W.moveToRangeStart, W.moveToRangeEnd);
        }

        // Deleted parts inside the moved sequences should be handled in a very special way.
        // Rather than marked as deleted in the source run, they should be moved to the target run
        // and then marked as deleted in the target run.
        private static List<ComparisonUnitAtom> AdjustMovedContentInTheComparisonUnitAtomList(
            List<ComparisonUnitAtom> comparisonUnitAtoms, 
            WmlComparerInternalSettings internalSettings
        )
        {
            XElement[] replaceAncestorElementsUpTo(IList<XElement> replaceIn, IList<XElement> replaceFrom, XName upTo)
            {
                return replaceFrom
                    .TakeWhile(e => {
                        return e.Name != upTo;
                    })
                    .Concat(replaceIn.SkipWhile(e => {
                        return e.Name != upTo;
                    }))
                    .ToArray();
            }

            // select move-from groups having deleted fragments
            var moveFromGroups = comparisonUnitAtoms
                .Where(ca => ca.MoveFromUnid != null)
                .GroupBy(ca => ca.MoveFromUnid)
                .Where(g => g.Any(ca => ca.MoveStatus == CorrelationStatus.Deleted))
                .ToList();
                
            // find corresponding move-to groups
            var moveToGroups = comparisonUnitAtoms
                .Where(ca => moveFromGroups.Any(g => g.Key == ca.MoveToUnid))
                .GroupBy(ca => ca.MoveToUnid)
                .ToList();

            foreach (var moveFromGroup in moveFromGroups)
            {
                var moveToGroup = moveToGroups.FirstOrDefault(g => g.Key == moveFromGroup.Key);
                if (moveToGroup == null)
                    continue;

                var moveFromGroupFragments = moveFromGroup
                    .GroupBy(ca => ca.MoveFragmentIndex)
                    .ToList();
                var moveToGroupFragments = moveToGroup
                    .GroupBy(ca => ca.MoveFragmentIndex)
                    .ToList();

                foreach (var moveFromGroupFragment in moveFromGroupFragments)
                {
                    // add items inside fragments should have the same move properties
                    var firstMoveFromGroupFragmentAtom = moveFromGroupFragment.FirstOrDefault();

                    // find deleted moved fragments in the move-from atoms and create their copies 
                    // at the appropriate position among move-to atoms in the common atom's list

                    if (firstMoveFromGroupFragmentAtom == null || firstMoveFromGroupFragmentAtom.MoveStatus != CorrelationStatus.Deleted)
                        continue;
                    
                    var moveFragmentIndex = moveFromGroupFragment.Key;

                    // for the deleted move-from fragment, there should no corresponding move-to fragment
                    // if such exists, it's an error
                    if (moveToGroupFragments.Any(g => g.Key == moveFragmentIndex))
                        throw new OpenXmlPowerToolsException("Internal error: unexpected move-to fragment for a deleted move-from fragment");

                    var siblingMoveToGroupFragmentAtom = (moveFragmentIndex == 0)
                        ? moveToGroupFragments.FirstOrDefault(g => g.Key == moveFragmentIndex + 1).FirstOrDefault()
                        : moveToGroupFragments.FirstOrDefault(g => g.Key == moveFragmentIndex - 1).LastOrDefault();
                    
                    if (siblingMoveToGroupFragmentAtom == null)
                        continue;

                    // clone and adjust ancestors 
                    var moveToGroupFragment = moveFromGroupFragment
                        .Select(ca => {
                            var nca = (ComparisonUnitAtom) ca.Clone();
                            nca.MoveToUnid = nca.MoveFromUnid;
                            nca.MoveFromUnid = null;
                            nca.CorrelationStatus = CorrelationStatus.Inserted;
                            nca.AncestorElements = replaceAncestorElementsUpTo(nca.AncestorElements, siblingMoveToGroupFragmentAtom.AncestorElements, W.r);
                            return nca;
                        })
                        .ToList();
                    
                    if (moveFragmentIndex == 0) {
                        // insert before the first atom of the sibling move-to fragment
                        comparisonUnitAtoms.InsertRange(
                            comparisonUnitAtoms.IndexOf(siblingMoveToGroupFragmentAtom),
                            moveToGroupFragment
                        );
                    } else {
                        // insert after the last atom of the sibling move-to fragment
                        comparisonUnitAtoms.InsertRange(
                            comparisonUnitAtoms.IndexOf(siblingMoveToGroupFragmentAtom) + 1,
                            moveToGroupFragment
                        );
                    }
                }
            }

            return comparisonUnitAtoms;
        }

    }
}