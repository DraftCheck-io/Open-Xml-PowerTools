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
            // TODO temporary disable 
            settings.DetectContentMoves = false;

            var originalWithUnids = PreProcessMarkup(original);
            WmlDocument merged = new WmlDocument(originalWithUnids);

            var revisedDocumentInfoListCount = revisedDocumentInfoList.Count();

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

                    foreach (var revisedDocumentInfo in revisedDocumentInfoList)
                    {
                        var internalSettings = new WmlComparerInternalSettings()
                        {
                            PreProcessMarkupInOriginal = false,
                            ResolveTrackingChanges = false, // do not wrap runs with w:ins and w:del 
                            IgnoreChangedContentDuringLcsAlgorithm = true, 
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
    }
}