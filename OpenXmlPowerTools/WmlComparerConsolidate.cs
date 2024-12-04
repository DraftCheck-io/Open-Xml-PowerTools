// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;

namespace OpenXmlPowerTools
{
    public class WmlComparerConsolidateSettings
    {
        public bool ConsolidateWithTable = true;
    }

    public static partial class WmlComparer
    {
        private class ConsolidationInfo
        {
            public string Revisor;
            public Color Color;
            public XElement RevisionElement;
            public bool InsertBefore = false;
            public string RevisionHash;
            public XElement[] Footnotes;
            public XElement[] Endnotes;
            public string RevisionString; // for debugging purposes only
        }

       /*****************************************************************************************************************/
        // Consolidate processes footnotes and endnotes in a particular fashion - if the unmodified document has a footnote
        // reference, and a delta has a footnote reference, we end up with two footnotes - one is unmodified, and is refered to
        // from the unmodified content.  The footnote reference in the delta refers to the modified footnote.  This is as it
        // should be.
        /*****************************************************************************************************************/
        public static WmlDocument Consolidate(
            WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings
        )
        {
            var consolidateSettings = new WmlComparerConsolidateSettings();
            return Consolidate(original, revisedDocumentInfoList, settings, consolidateSettings);
        }

        public static WmlDocument Consolidate(
            WmlDocument original,
            List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
            WmlComparerSettings settings, 
            WmlComparerConsolidateSettings consolidateSettings
        )
        {
            // DraftCheck : temporary disable 
            settings.DetectContentMoves = false;

            var internalSettings = new WmlComparerInternalSettings()
            {
                StartingIdForFootnotesEndnotes = 3000
            };

#if false
            var now = DateTime.Now;
            var tempName = String.Format("{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", now.Year - 2000, now.Month, now.Day, now.Hour, now.Minute, now.Second);
            FileInfo fi = new FileInfo("./WmlComparer.Consolidate-" + tempName + "-Original.docx");
            File.WriteAllBytes(fi.FullName, original.DocumentByteArray);
            for (int i = 0; i < revisedDocumentInfoList.Count(); i++)
            {
                fi = new FileInfo("./WmlComparer.Consolidate-" + tempName + string.Format("-Revised-{0}", i) + ".docx");
                File.WriteAllBytes(fi.FullName, revisedDocumentInfoList.ElementAt(i).RevisedDocument.DocumentByteArray);
            }
            StringBuilder sbt = new StringBuilder();
            int count = 0;
            foreach (var rev in revisedDocumentInfoList)
            {
                sbt.Append("Revised #" + (count++).ToString() + Environment.NewLine);
                sbt.Append("Color:" + rev.Color.ToString() + Environment.NewLine);
                sbt.Append("Revisor:" + rev.Revisor + Environment.NewLine);
                sbt.Append("" + Environment.NewLine);
            }
            sbt.Append("settings.AuthorForRevisions:" + settings.AuthorForRevisions + Environment.NewLine);
            sbt.Append("settings.CaseInsensitive:" + settings.CaseInsensitive.ToString() + Environment.NewLine);
            sbt.Append("settings.CultureInfo:" + settings.CultureInfo.ToString() + Environment.NewLine);
            sbt.Append("settings.DateTimeForRevisions:" + settings.DateTimeForRevisions.ToString() + Environment.NewLine);
            sbt.Append("settings.DetailThreshold:" + settings.DetailThreshold.ToString() + Environment.NewLine);
            sbt.Append("settings.StartingIdForFootnotesEndnotes:" + inter.StartingIdForFootnotesEndnotes.ToString() + Environment.NewLine);
            sbt.Append("settings.WordSeparators:" + settings.WordSeparators.Select(ws => ws.ToString()).StringConcatenate() + Environment.NewLine);
            //sb.Append(":" + settings);
            fi = new FileInfo("./WmlComparer.Consolidate-" + tempName + "-Settings.txt");
            File.WriteAllText(fi.FullName, sbt.ToString());
#endif

            // pre-process the original, so that it already has unids for all elements
            // then when comparing all documents to the original, each one will have the unid as appropriate
            // for all revision block-level content
            //   set unid to look for
            //   while true
            //     determine where to insert
            //       get the unid for the revision
            //       look it up in the original.  if find it, then insert after that element
            //       if not in the original
            //         look backwards in revised document, set unid to look for, do the loop again
            //       if get to the beginning of the document
            //         insert at beginning of document

            var originalWithUnids = PreProcessMarkup(original, internalSettings.StartingIdForFootnotesEndnotes);
            WmlDocument consolidated = new WmlDocument(originalWithUnids);

            if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var name1 = "Original-with-Unids.docx";
                var preProcFi1 = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                originalWithUnids.SaveAs(preProcFi1.FullName);
            }

            var revisedDocumentInfoListCount = revisedDocumentInfoList.Count();

            using (MemoryStream consolidatedMs = new MemoryStream())
            {
                consolidatedMs.Write(consolidated.DocumentByteArray, 0, consolidated.DocumentByteArray.Length);
                using (WordprocessingDocument consolidatedWDoc = WordprocessingDocument.Open(consolidatedMs, true))
                {
                    var consolidatedMainDocPart = consolidatedWDoc.MainDocumentPart;
                    var consolidatedMainDocPartXDoc = consolidatedMainDocPart.GetXDocument();

                    // save away last sectPr
                    XElement savedSectPr = consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Elements(W.sectPr)
                        .LastOrDefault();
                    consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Elements(W.sectPr)
                        .Remove();

                    var consolidatedByUnid = consolidatedMainDocPartXDoc
                        .Descendants()
                        .Where(d => (d.Name == W.p || d.Name == W.tbl) && d.Attribute(PtOpenXml.Unid) != null)
                        .ToDictionary(d => (string)d.Attribute(PtOpenXml.Unid));

                    int deltaNbr = 1;
                    foreach (var revisedDocumentInfo in revisedDocumentInfoList)
                    {
                        internalSettings.StartingIdForFootnotesEndnotes = (deltaNbr * 2000) + 3000;
                        var delta = WmlComparer.CompareInternal(originalWithUnids, revisedDocumentInfo.RevisedDocument, settings, internalSettings);

                        if (s_SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
                        {
                            var name1 = string.Format("Delta-{0}.docx", deltaNbr++);
                            var deltaFi = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name1));
                            delta.SaveAs(deltaFi.FullName);
                        }

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
                                var blockLevelContentToMove = deltaMainDocPartXDoc
                                    .Root
                                    .DescendantsTrimmed(d => d.Name == W.txbxContent || d.Name == W.tr)
                                    .Where(d => d.Name == W.p || d.Name == W.tbl)
                                    .Where(d => d.Descendants().Any(z => z.Name == W.ins || z.Name == W.del) ||
                                        ContentContainsFootnoteEndnoteReferencesThatHaveRevisions(d, wDocDelta))
                                    .ToList();

                                foreach (var revision in blockLevelContentToMove)
                                {
                                    var elementLookingAt = revision;
                                    while (true)
                                    {
                                        var unid = (string)elementLookingAt.Attribute(PtOpenXml.Unid);
                                        if (unid == null)
                                            throw new OpenXmlPowerToolsException("Internal error");

                                        XElement elementToInsertAfter = null;
                                        if (consolidatedByUnid.ContainsKey(unid))
                                            elementToInsertAfter = consolidatedByUnid[unid];

                                        if (elementToInsertAfter != null)
                                        {
                                            ConsolidationInfo ci = new ConsolidationInfo();
                                            ci.Revisor = revisedDocumentInfo.Revisor;
                                            ci.Color = revisedDocumentInfo.Color;
                                            ci.RevisionElement = revision;
                                            ci.Footnotes = revision
                                                .Descendants(W.footnoteReference)
                                                .Select(fr =>
                                                {
                                                    var id = (int)fr.Attribute(W.id);
                                                    var fnXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
                                                    var footnote = fnXDoc.Root.Elements(W.footnote).FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                                    if (footnote == null)
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    return footnote;
                                                })
                                                .ToArray();
                                            ci.Endnotes = revision
                                                .Descendants(W.endnoteReference)
                                                .Select(er =>
                                                {
                                                    var id = (int)er.Attribute(W.id);
                                                    var enXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
                                                    var endnote = enXDoc.Root.Elements(W.endnote).FirstOrDefault(en => (int)en.Attribute(W.id) == id);
                                                    if (endnote == null)
                                                        throw new OpenXmlPowerToolsException("Internal Error");
                                                    return endnote;
                                                })
                                                .ToArray();
                                            AddToAnnotation(
                                                wDocDelta,
                                                consolidatedWDoc,
                                                elementToInsertAfter,
                                                ci,
                                                settings,
                                                internalSettings
                                            );
                                            break;
                                        }
                                        else
                                        {
                                            // find an element to insert after
                                            var elementBeforeRevision = elementLookingAt
                                                .SiblingsBeforeSelfReverseDocumentOrder()
                                                .FirstOrDefault(e => e.Attribute(PtOpenXml.Unid) != null);
                                            if (elementBeforeRevision == null)
                                            {
                                                var firstElement = consolidatedMainDocPartXDoc
                                                    .Root
                                                    .Element(W.body)
                                                    .Elements()
                                                    .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);

                                                ConsolidationInfo ci = new ConsolidationInfo();
                                                ci.Revisor = revisedDocumentInfo.Revisor;
                                                ci.Color = revisedDocumentInfo.Color;
                                                ci.RevisionElement = revision;
                                                ci.InsertBefore = true;
                                                ci.Footnotes = revision
                                                    .Descendants(W.footnoteReference)
                                                    .Select(fr =>
                                                    {
                                                        var id = (int)fr.Attribute(W.id);
                                                        var fnXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
                                                        var footnote = fnXDoc.Root.Elements(W.footnote).FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                                        if (footnote == null)
                                                            throw new OpenXmlPowerToolsException("Internal Error");
                                                        return footnote;
                                                    })
                                                    .ToArray();
                                                ci.Endnotes = revision
                                                    .Descendants(W.endnoteReference)
                                                    .Select(er =>
                                                    {
                                                        var id = (int)er.Attribute(W.id);
                                                        var enXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
                                                        var endnote = enXDoc.Root.Elements(W.endnote).FirstOrDefault(en => (int)en.Attribute(W.id) == id);
                                                        if (endnote == null)
                                                            throw new OpenXmlPowerToolsException("Internal Error");
                                                        return endnote;
                                                    })
                                                    .ToArray();
                                                AddToAnnotation(
                                                    wDocDelta,
                                                    consolidatedWDoc,
                                                    firstElement,
                                                    ci,
                                                    settings,
                                                    internalSettings
                                                );
                                                break;
                                            }
                                            else
                                            {
                                                elementLookingAt = elementBeforeRevision;
                                                continue;
                                            }
                                        }
                                    }
                                }
                                CopyMissingStylesFromOneDocToAnother(wDocDelta, consolidatedWDoc);
                            }
                        }
                    }

                    // at this point, everything is added as an annotation, from all documents to be merged.
                    // so now the process is to go through and add the annotations to the document
                    var elementsToProcess = consolidatedMainDocPartXDoc
                        .Root
                        .Descendants()
                        .Where(d => d.Annotation<List<ConsolidationInfo>>() != null)
                        .ToList();

                    var emptyParagraph = new XElement(W.p,
                        new XElement(W.pPr,
                            new XElement(W.spacing,
                                new XAttribute(W.after, "0"),
                                new XAttribute(W.line, "240"),
                                new XAttribute(W.lineRule, "auto"))));

                    foreach (var ele in elementsToProcess)
                    {
                        var lci = ele.Annotation<List<ConsolidationInfo>>();

                        // process before
                        var contentToAddBefore = lci
                            .Where(ci => ci.InsertBefore == true)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        ele.AddBeforeSelf(contentToAddBefore);

                        // process after
                        // if all revisions from all revisors are exactly the same, then instead of adding multiple tables after
                        // that contains the revisions, then simply replace the paragraph with the one with the revisions.
                        // RC004 documents contain the test data to exercise this.

                        var lciCount = lci.Where(ci => ci.InsertBefore == false).Count();

                        if (lciCount > 1 && lciCount == revisedDocumentInfoListCount)
                        {
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            // This is the code that determines if revisions should be consolidated into one.

                            var uniqueRevisions = lci
                                .Where(ci => ci.InsertBefore == false)
                                .GroupBy(ci =>
                                {
                                    // Get a hash after first accepting revisions and compressing the text.
                                    var acceptedRevisionElement = RevisionProcessor.AcceptRevisionsForElement(ci.RevisionElement);
                                    var sha1Hash = PtUtils.SHA1HashStringForUTF8String(acceptedRevisionElement.Value.Replace(" ", "").Replace(" ", "").Replace(" ", "").Replace("\n", "").Replace(".", "").Replace(",", "").ToUpper());
                                    return sha1Hash;
                                })
                                .OrderByDescending(g => g.Count())
                                .ToList();
                            var uniqueRevisionCount = uniqueRevisions.Count();

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            if (uniqueRevisionCount == 1)
                            {
                                MoveFootnotesEndnotesForConsolidatedRevisions(lci.First(), consolidatedWDoc);

                                var dummyElement = new XElement("dummy", lci.First().RevisionElement);

                                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                                {
                                    var aut = rev.Attribute(W.author);
                                    aut.Value = "ITU";
                                }

                                ele.ReplaceWith(dummyElement.Elements());
                                continue;
                            }

                            // this is the location where we have determined that there are the same number of revisions for this paragraph as there are revision documents.
                            // however, the hash for all of them were not the same.
                            // therefore, they would be added to the consolidated document as separate revisions.

                            // create a log that shows what is different, in detail.
                            if (settings.LogCallback != null)
                            {
                                StringBuilder sb = new StringBuilder();
                                sb.Append("====================================================================================================" + nl);
                                sb.Append("Non-Consolidated Revision" + nl);
                                sb.Append("====================================================================================================" + nl);
                                foreach (var urList in uniqueRevisions)
                                {
                                    var revisorList = urList.Select(ur => ur.Revisor + " : ").StringConcatenate().TrimEnd(' ', ':');
                                    sb.Append("Revisors: " + revisorList + nl);
                                    var str = RevisionToLogFormTransform(urList.First().RevisionElement, 0, false);
                                    sb.Append(str);
                                    sb.Append("=========================" + nl);
                                }
                                sb.Append(nl);
                                settings.LogCallback(sb.ToString());
                            }
                        }

                        // todo this is where it assembles the content to put into a single cell table
                        // the magic function is AssembledConjoinedRevisionContent

                        var contentToAddAfter = lci
                            .Where(ci => ci.InsertBefore == false)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        ele.AddAfterSelf(contentToAddAfter);
                    }

#if false
                    // old code
                    foreach (var ele in elementsToProcess)
                    {
                        var lci = ele.Annotation<List<ConsolidationInfo>>();

                        // if all revisions from all revisors are exactly the same, then instead of adding multiple tables after
                        // that contains the revisions, then simply replace the paragraph with the one with the revisions.
                        // RC004 documents contain the test data to exercise this.

                        var lciCount = lci.Count();

                        if (lci.Count() > 1 && lciCount == revisedDocumentInfoListCount)
                        {
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            // This is the code that determines if revisions should be consolidated into one.

                            var uniqueRevisions = lci
                                .GroupBy(ci =>
                                {
                                    // Get a hash after first accepting revisions and compressing the text.
                                    var ciz = ci;

                                    var acceptedRevisionElement = RevisionProcessor.AcceptRevisionsForElement(ci.RevisionElement);
                                    var text = acceptedRevisionElement.Value
                                        .Replace(" ", "")
                                        .Replace(" ", "")
                                        .Replace(" ", "")
                                        .Replace("\n", "");
                                    var sha1Hash = PtUtils.SHA1HashStringForUTF8String(text);
                                    return ci.InsertBefore.ToString() + sha1Hash;
                                })
                                .OrderByDescending(g => g.Count())
                                .ToList();
                            var uniqueRevisionCount = uniqueRevisions.Count();

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            if (uniqueRevisionCount == 1)
                            {
                                MoveFootnotesEndnotesForConsolidatedRevisions(lci.First(), consolidatedWDoc);

                                var dummyElement = new XElement("dummy", lci.First().RevisionElement);

                                foreach(var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                                {
                                    var aut = rev.Attribute(W.author);
                                    aut.Value = "ITU";
                                }

                                ele.ReplaceWith(dummyElement.Elements());
                                continue;
                            }

                            // this is the location where we have determined that there are the same number of revisions for this paragraph as there are revision documents.
                            // however, the hash for all of them were not the same.
                            // therefore, they would be added to the consolidated document as separate revisions.

                            // create a log that shows what is different, in detail.
                            if (settings.LogCallback != null)
                            {
                                StringBuilder sb = new StringBuilder();
                                sb.Append("====================================================================================================" + nl);
                                sb.Append("Non-Consolidated Revision" + nl);
                                sb.Append("====================================================================================================" + nl);
                                foreach (var urList in uniqueRevisions)
                                {
                                    var revisorList = urList.Select(ur => ur.Revisor + " : ").StringConcatenate().TrimEnd(' ', ':');
                                    sb.Append("Revisors: " + revisorList + nl);
                                    var str = RevisionToLogFormTransform(urList.First().RevisionElement, 0, false);
                                    sb.Append(str);
                                    sb.Append("=========================" + nl);
                                }
                                sb.Append(nl);
                                settings.LogCallback(sb.ToString());
                            }
                        }

                        var contentToAddBefore = lci
                            .Where(ci => ci.InsertBefore == true)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        var contentToAddAfter = lci
                            .Where(ci => ci.InsertBefore == false)
                            .GroupAdjacent(ci => ci.Revisor + ci.Color.ToString())
                            .Select((groupedCi, idx) => AssembledConjoinedRevisionContent(emptyParagraph, groupedCi, idx, consolidatedWDoc, consolidateSettings));
                        ele.AddBeforeSelf(contentToAddBefore);
                        ele.AddAfterSelf(contentToAddAfter);
                    }
#endif

                    consolidatedMainDocPartXDoc
                        .Root
                        .Element(W.body)
                        .Add(savedSectPr);

                    AddTableGridStyleToStylesPart(consolidatedWDoc.MainDocumentPart.StyleDefinitionsPart);
                    FixUpRevisionIds(consolidatedWDoc, consolidatedMainDocPartXDoc);
                    IgnorePt14NamespaceForFootnotesEndnotes(consolidatedWDoc);
                    FixUpDocPrIds(consolidatedWDoc);
                    FixUpShapeIds(consolidatedWDoc);
                    FixUpGroupIds(consolidatedWDoc);
                    FixUpShapeTypeIds(consolidatedWDoc);
                    RemoveCustomMarkFollows(consolidatedWDoc);
                    WmlComparer.IgnorePt14Namespace(consolidatedMainDocPartXDoc.Root);
                    consolidatedWDoc.MainDocumentPart.PutXDocument();
                    AddFootnotesEndnotesStyles(consolidatedWDoc);
                }

                var newConsolidatedDocument = new WmlDocument("consolidated.docx", consolidatedMs.ToArray());
                return newConsolidatedDocument;
            }
        }

        private static void RemoveCustomMarkFollows(WordprocessingDocument consolidatedWDoc)
        {
            var mxDoc = consolidatedWDoc.MainDocumentPart.GetXDocument();
            mxDoc.Root.Descendants().Attributes(W.customMarkFollows).Remove();
            consolidatedWDoc.MainDocumentPart.PutXDocument();
        }

        private static void MoveFootnotesEndnotesForConsolidatedRevisions(ConsolidationInfo ci, WordprocessingDocument wDocConsolidated)
        {
            var consolidatedFootnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
            var consolidatedEndnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();

            int maxFootnoteId = 1;
            if (consolidatedFootnoteXDoc.Root.Elements(W.footnote).Any())
                maxFootnoteId = consolidatedFootnoteXDoc.Root.Elements(W.footnote).Select(e => (int)e.Attribute(W.id)).Max();
            int maxEndnoteId = 1;
            if (consolidatedEndnoteXDoc.Root.Elements(W.endnote).Any())
                maxEndnoteId = consolidatedEndnoteXDoc.Root.Elements(W.endnote).Select(e => (int)e.Attribute(W.id)).Max(); ;

            /// At this point, content might contain a footnote or endnote reference.
            /// Need to add the footnote / endnote into the consolidated document (with the same guid id)
            /// Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

            if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
            {
                var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                {
                    var id = (int)footnoteReference.Attribute(W.id);
                    var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    var newId = maxFootnoteId + 1;
                    maxFootnoteId++;
                    footnoteReference.Attribute(W.id).Value = newId.ToString();
                    var clonedFootnote = new XElement(footnote);
                    clonedFootnote.Attribute(W.id).Value = newId.ToString();
                    footnoteXDoc.Root.Add(clonedFootnote);
                }
                wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
            }

            if (ci.RevisionElement.Descendants(W.endnoteReference).Any())
            {
                var endnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();
                foreach (var endnoteReference in ci.RevisionElement.Descendants(W.endnoteReference))
                {
                    var id = (int)endnoteReference.Attribute(W.id);
                    var endnote = ci.Endnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                    var newId = maxEndnoteId + 1;
                    maxEndnoteId++;
                    endnoteReference.Attribute(W.id).Value = newId.ToString();
                    var clonedEndnote = new XElement(endnote);
                    clonedEndnote.Attribute(W.id).Value = newId.ToString();
                    endnoteXDoc.Root.Add(clonedEndnote);
                }
                wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
            }
        }

        private static bool ContentContainsFootnoteEndnoteReferencesThatHaveRevisions(XElement element, WordprocessingDocument wDocDelta)
        {
            var footnoteEndnoteReferences = element.Descendants().Where(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference);
            if (!footnoteEndnoteReferences.Any())
                return false;
            var footnoteXDoc = wDocDelta.MainDocumentPart.FootnotesPart.GetXDocument();
            var endnoteXDoc = wDocDelta.MainDocumentPart.EndnotesPart.GetXDocument();
            foreach (var note in footnoteEndnoteReferences)
            {
                XElement fnen = null;
                if (note.Name == W.footnoteReference)
                {
                    var id = (int)note.Attribute(W.id);
                    fnen = footnoteXDoc
                        .Root
                        .Elements(W.footnote)
                        .FirstOrDefault(n => (int)n.Attribute(W.id) == id);
                    if (fnen.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Any())
                        return true;
                }
                if (note.Name == W.endnoteReference)
                {
                    var id = (int)note.Attribute(W.id);
                    fnen = endnoteXDoc
                        .Root
                        .Elements(W.endnote)
                        .FirstOrDefault(n => (int)n.Attribute(W.id) == id);
                    if (fnen.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Any())
                        return true;
                }
            }
            return false;
        }

        private static XElement[] AssembledConjoinedRevisionContent(XElement emptyParagraph, IGrouping<string, ConsolidationInfo> groupedCi, int idx, WordprocessingDocument wDocConsolidated,
            WmlComparerConsolidateSettings consolidateSettings)
        {
            var consolidatedFootnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
            var consolidatedEndnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();

            int maxFootnoteId = 1;
            if (consolidatedFootnoteXDoc.Root.Elements(W.footnote).Any())
                maxFootnoteId = consolidatedFootnoteXDoc.Root.Elements(W.footnote).Select(e => (int)e.Attribute(W.id)).Max();
            int maxEndnoteId = 1;
            if (consolidatedEndnoteXDoc.Root.Elements(W.endnote).Any())
                maxEndnoteId = consolidatedEndnoteXDoc.Root.Elements(W.endnote).Select(e => (int)e.Attribute(W.id)).Max(); ;

            var revisor = groupedCi.First().Revisor;

            var captionParagraph = new XElement(W.p,
                new XElement(W.pPr,
                    new XElement(W.jc, new XAttribute(W.val, "both")),
                    new XElement(W.rPr,
                        new XElement(W.b),
                        new XElement(W.bCs))),
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.b),
                        new XElement(W.bCs)),
                    new XElement(W.t, revisor)));

            var colorRgb = groupedCi.First().Color.ToArgb();
            var colorString = colorRgb.ToString("X");
            if (colorString.Length == 8)
                colorString = colorString.Substring(2);

            if (consolidateSettings.ConsolidateWithTable)
            {
                var table = new XElement(W.tbl,
                    new XElement(W.tblPr,
                        new XElement(W.tblStyle, new XAttribute(W.val, "TableGridForRevisions")),
                        new XElement(W.tblW,
                            new XAttribute(W._w, "0"),
                            new XAttribute(W.type, "auto")),
                        new XElement(W.shd,
                            new XAttribute(W.val, "clear"),
                            new XAttribute(W.color, "auto"),
                            new XAttribute(W.fill, colorString)),
                        new XElement(W.tblLook,
                            new XAttribute(W.firstRow, "0"),
                            new XAttribute(W.lastRow, "0"),
                            new XAttribute(W.firstColumn, "0"),
                            new XAttribute(W.lastColumn, "0"),
                            new XAttribute(W.noHBand, "0"),
                            new XAttribute(W.noVBand, "0"))),
                    new XElement(W.tblGrid,
                        new XElement(W.gridCol, new XAttribute(W._w, "9576"))),
                    new XElement(W.tr,
                        new XElement(W.tc,
                            new XElement(W.tcPr,
                            new XElement(W.shd,
                                new XAttribute(W.val, "clear"),
                                new XAttribute(W.color, "auto"),
                                new XAttribute(W.fill, colorString))),
                            captionParagraph,
                            groupedCi.Select(ci =>
                            {
                                /// At this point, content might contain a footnote or endnote reference.
                                /// Need to add the footnote / endnote into the consolidated document (with the same guid id)
                                /// Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

                                if (ci.RevisionElement.Descendants(W.endnoteReference).Any())
                                {
                                    var endnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();
                                    foreach (var endnoteReference in ci.RevisionElement.Descendants(W.endnoteReference))
                                    {
                                        var id = (int)endnoteReference.Attribute(W.id);
                                        var endnote = ci.Endnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                        var newId = maxEndnoteId + 1;
                                        maxEndnoteId++;
                                        endnoteReference.Attribute(W.id).Value = newId.ToString();
                                        var clonedEndnote = new XElement(endnote);
                                        clonedEndnote.Attribute(W.id).Value = newId.ToString();
                                        endnoteXDoc.Root.Add(clonedEndnote);
                                    }
                                    wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
                                }

                                if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
                                {
                                    var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                                    foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                                    {
                                        var id = (int)footnoteReference.Attribute(W.id);
                                        var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                                        var newId = maxFootnoteId + 1;
                                        maxFootnoteId++;
                                        footnoteReference.Attribute(W.id).Value = newId.ToString();
                                        var clonedFootnote = new XElement(footnote);
                                        clonedFootnote.Attribute(W.id).Value = newId.ToString();
                                        footnoteXDoc.Root.Add(clonedFootnote);
                                    }
                                    wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
                                }

                                // it is important that this code follows the code above, because the code above updates ci.RevisionElement (using DML)

                                XElement paraAfter = null;
                                if (ci.RevisionElement.Name == W.tbl)
                                    paraAfter = emptyParagraph;
                                var revisionInTable = new[] {
                                    ci.RevisionElement,
                                    paraAfter,
                                    };

                                return revisionInTable;
                            }))));

                // if the last paragraph has a deleted paragraph mark, then remove the deletion from the paragraph mark.  This is to prevent Word from misbehaving.
                // the last paragraph in a cell must not have a deleted paragraph mark.
                var theCell = table
                    .Descendants(W.tc)
                    .FirstOrDefault();
                var lastPara = theCell
                    .Elements(W.p)
                    .LastOrDefault();
                if (lastPara != null)
                {
                    var isDeleted = lastPara
                        .Elements(W.pPr)
                        .Elements(W.rPr)
                        .Elements(W.del)
                        .Any();
                    if (isDeleted)
                        lastPara
                            .Elements(W.pPr)
                            .Elements(W.rPr)
                            .Elements(W.del)
                            .Remove();
                }

                var content = new[] {
                                    idx == 0 ? emptyParagraph : null,
                                    table,
                                    emptyParagraph,
                                };

                var dummyElement = new XElement("dummy", content);

                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                {
                    var aut = rev.Attribute(W.author);
                    aut.Value = revisor;
                }

                return dummyElement.Elements().ToArray();
            }
            else
            {
                var content = groupedCi.Select(ci =>
                {
                    XElement paraAfter = null;
                    if (ci.RevisionElement.Name == W.tbl)
                        paraAfter = emptyParagraph;
                    var revisionInTable = new[] {
                                    ci.RevisionElement,
                                    paraAfter,
                                    };

                    /// At this point, content might contain a footnote or endnote reference.
                    /// Need to add the footnote / endnote into the consolidated document (with the same guid id)
                    /// Because of preprocessing of the documents, all footnote and endnote references will be unique at this point

                    if (ci.RevisionElement.Descendants(W.footnoteReference).Any())
                    {
                        var footnoteXDoc = wDocConsolidated.MainDocumentPart.FootnotesPart.GetXDocument();
                        foreach (var footnoteReference in ci.RevisionElement.Descendants(W.footnoteReference))
                        {
                            var id = (int)footnoteReference.Attribute(W.id);
                            var footnote = ci.Footnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                            var newId = maxFootnoteId + 1;
                            maxFootnoteId++;
                            footnoteReference.Attribute(W.id).Value = newId.ToString();
                            var clonedFootnote = new XElement(footnote);
                            clonedFootnote.Attribute(W.id).Value = newId.ToString();
                            footnoteXDoc.Root.Add(clonedFootnote);
                        }
                        wDocConsolidated.MainDocumentPart.FootnotesPart.PutXDocument();
                    }

                    if (ci.RevisionElement.Descendants(W.endnoteReference).Any())
                    {
                        var endnoteXDoc = wDocConsolidated.MainDocumentPart.EndnotesPart.GetXDocument();
                        foreach (var endnoteReference in ci.RevisionElement.Descendants(W.endnoteReference))
                        {
                            var id = (int)endnoteReference.Attribute(W.id);
                            var endnote = ci.Endnotes.FirstOrDefault(fn => (int)fn.Attribute(W.id) == id);
                            var newId = maxEndnoteId + 1;
                            maxEndnoteId++;
                            endnoteReference.Attribute(W.id).Value = newId.ToString();
                            var clonedEndnote = new XElement(endnote);
                            clonedEndnote.Attribute(W.id).Value = newId.ToString();
                            endnoteXDoc.Root.Add(clonedEndnote);
                        }
                        wDocConsolidated.MainDocumentPart.EndnotesPart.PutXDocument();
                    }

                    return revisionInTable;
                });

                var dummyElement = new XElement("dummy",
                    content.SelectMany(m => m));

                foreach (var rev in dummyElement.Descendants().Where(d => d.Attribute(W.author) != null))
                {
                    var aut = rev.Attribute(W.author);
                    aut.Value = revisor;
                }

                return dummyElement.Elements().ToArray();
            }
        }

        private static void AddToAnnotation(
            WordprocessingDocument wDocDelta,
            WordprocessingDocument consolidatedWDoc,
            XElement elementToInsertAfter,
            ConsolidationInfo consolidationInfo,
            WmlComparerSettings settings,
            WmlComparerInternalSettings internalSettings
        )
        {
            Package packageOfDeletedContent = wDocDelta.MainDocumentPart.OpenXmlPackage.Package;
            Package packageOfNewContent = consolidatedWDoc.MainDocumentPart.OpenXmlPackage.Package;
            PackagePart partInDeletedDocument = packageOfDeletedContent.GetPart(wDocDelta.MainDocumentPart.Uri);
            PackagePart partInNewDocument = packageOfNewContent.GetPart(consolidatedWDoc.MainDocumentPart.Uri);
            consolidationInfo.RevisionElement = MoveRelatedPartsToDestination(partInDeletedDocument, partInNewDocument, consolidationInfo.RevisionElement);

            var clonedForHashing = (XElement)CloneBlockLevelContentForHashing(consolidatedWDoc.MainDocumentPart, consolidationInfo.RevisionElement, false, settings, internalSettings);
            clonedForHashing.Descendants().Where(d => d.Name == W.ins || d.Name == W.del).Attributes(W.id).Remove();
            var shaString = clonedForHashing.ToString(SaveOptions.DisableFormatting)
                .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
            var sha1Hash = PtUtils.SHA1HashStringForUTF8String(shaString);
            consolidationInfo.RevisionString = shaString;
            consolidationInfo.RevisionHash = sha1Hash;

            var annotationList = elementToInsertAfter.Annotation<List<ConsolidationInfo>>();
            if (annotationList == null)
            {
                annotationList = new List<ConsolidationInfo>();
                elementToInsertAfter.AddAnnotation(annotationList);
            }
            annotationList.Add(consolidationInfo);
        }

        private static void AddTableGridStyleToStylesPart(StyleDefinitionsPart styleDefinitionsPart)
        {
            var sXDoc = styleDefinitionsPart.GetXDocument();
            var tableGridStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "TableGridForRevisions");
            if (tableGridStyle == null)
            {
                var tableGridForRevisionsStyleMarkup =
@"<w:style w:type=""table""
         w:styleId=""TableGridForRevisions""
         xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:name w:val=""Table Grid For Revisions""/>
  <w:basedOn w:val=""TableNormal""/>
  <w:rsid w:val=""0092121A""/>
  <w:rPr>
    <w:rFonts w:asciiTheme=""minorHAnsi""
              w:eastAsiaTheme=""minorEastAsia""
              w:hAnsiTheme=""minorHAnsi""
              w:cstheme=""minorBidi""/>
    <w:sz w:val=""22""/>
    <w:szCs w:val=""22""/>
  </w:rPr>
  <w:tblPr>
    <w:tblBorders>
      <w:top w:val=""single""
             w:sz=""4""
             w:space=""0""
             w:color=""auto""/>
      <w:left w:val=""single""
              w:sz=""4""
              w:space=""0""
              w:color=""auto""/>
      <w:bottom w:val=""single""
                w:sz=""4""
                w:space=""0""
                w:color=""auto""/>
      <w:right w:val=""single""
               w:sz=""4""
               w:space=""0""
               w:color=""auto""/>
      <w:insideH w:val=""single""
                 w:sz=""4""
                 w:space=""0""
                 w:color=""auto""/>
      <w:insideV w:val=""single""
                 w:sz=""4""
                 w:space=""0""
                 w:color=""auto""/>
    </w:tblBorders>
  </w:tblPr>
</w:style>";
                var tgsElement = XElement.Parse(tableGridForRevisionsStyleMarkup);
                sXDoc.Root.Add(tgsElement);
            }
            var tableNormalStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(s => (string)s.Attribute(W.styleId) == "TableNormal");
            if (tableNormalStyle == null)
            {
                var tableNormalStyleMarkup =
@"<w:style w:type=""table""
           w:default=""1""
           w:styleId=""TableNormal""
           xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:name w:val=""Normal Table""/>
    <w:uiPriority w:val=""99""/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w=""0""
                w:type=""dxa""/>
      <w:tblCellMar>
        <w:top w:w=""0""
               w:type=""dxa""/>
        <w:left w:w=""108""
                w:type=""dxa""/>
        <w:bottom w:w=""0""
                  w:type=""dxa""/>
        <w:right w:w=""108""
                 w:type=""dxa""/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>";
                var tnsElement = XElement.Parse(tableNormalStyleMarkup);
                sXDoc.Root.Add(tnsElement);
            }
            styleDefinitionsPart.PutXDocument();
        }
        
    }
}