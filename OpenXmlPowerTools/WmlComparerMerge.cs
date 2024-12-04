using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class WmlComparerMergeSettings
    {
        //
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

                        var internalSettings = new WmlComparerInternalSettings()
                        {
                            PreProcessMarkupInOriginal = false,
                            MergeMode = true, 
                            MergeIteration = i,
                            // for last item, we need to resolve all accumulated tracking changes
                            ResolveTrackingChanges = isLast,
                            MergeRevisors = revisors,
                        };

                        var revised = revisedDocumentInfo.RevisedDocument;
                        var delta = WmlComparer.CompareInternal(merged, revised, settings, internalSettings);

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

        private static void StoreChangeTrackingStatusesForMerge(XDocument doc, int mergeIteration)
        {
            doc.Root
                .Descendants()
                .Where(e => e.Attribute(PtOpenXml.Status) != null && e.Attribute(PtOpenXml.MergeStatus) == null)
                .ToList()
                .ForEach(e => {
                    e.SetAttributeValue(PtOpenXml.MergeStatus, e.Attribute(PtOpenXml.Status).Value);
                    e.SetAttributeValue(PtOpenXml.MergeIteration, mergeIteration);
                });
        }

    }
}