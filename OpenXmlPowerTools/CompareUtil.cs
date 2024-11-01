using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OpenXmlPowerTools
{
    public static class SimilarityUtil
    {
        private static bool AreSequencesEqual<T>(T[] sequence1, T[] sequence2)
        {
            if (sequence1.Length != sequence2.Length)
            {
                return false;
            }

            for (int i = 0; i < sequence1.Length; i++)
            {
                if (!EqualityComparer<T>.Default.Equals(sequence1[i], sequence2[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static char[] PrepareStringForComparison(string str)
        {
            return StringUtil.TrimWhiteSpaces(str).ToLower().ToCharArray();
        }

        private static int CalcLevenshteinEditDistance<T>(T[] sequence1, T[] sequence2)
        {
            var costs = new int[sequence2.Length + 1];
            for (int i = 0; i <= sequence1.Length; i++)
            {
                int lastValue = i;
                for (int j = 0; j <= sequence2.Length; j++)
                {
                    if (i == 0)
                    {
                        costs[j] = j;
                    }
                    else if (j > 0)
                    {
                        int newValue = costs[j - 1];
                        if (!EqualityComparer<T>.Default.Equals(sequence1[i - 1], sequence2[j - 1]))
                        {
                            newValue = Math.Min(Math.Min(newValue, lastValue), costs[j]) + 1;
                        }
                        costs[j - 1] = lastValue;
                        lastValue = newValue;
                    }
                }
                if (i > 0)
                {
                    costs[sequence2.Length] = lastValue;
                }
            }
            return costs[sequence2.Length];
        }

        private static double CalcSimilarity<T>(T[] sequence1, T[] sequence2)
        {
            if (AreSequencesEqual(sequence1, sequence2))
            {
                return 1;
            }

            var longerSequence = sequence1;
            var shorterSequence = sequence2;

            if (sequence1.Length < sequence2.Length)
            {
                longerSequence = sequence2;
                shorterSequence = sequence1;
            }

            int longerSequenceLength = longerSequence.Length;
            if (longerSequenceLength == 0)
            {
                return 1;
            }

            return (longerSequenceLength - CalcLevenshteinEditDistance(longerSequence, shorterSequence)) / (double)longerSequenceLength;
        }

        public static double[,] CalcSimilarityMatrix<T>(T[][] sequences1, T[][] sequences2)
        {
            var matrix = new double[sequences1.Length, sequences2.Length];

            for (int i = 0; i < sequences1.Length; i++)
            // Parallel.For(0, sequences1.Length, i =>
            {
                T[] sequence1 = sequences1[i];
                for (int j = 0; j < sequences2.Length; j++)
                {
                    T[] sequence2 = sequences2[j];
                    matrix[i, j] = (sequence1.Length > 0 && sequence2.Length > 0) ? CalcSimilarity(sequence1, sequence2) : 0;
                }
            };

            return matrix;
        }

        public static double[,] CalcSimilarityMatrix(string[] sequences1, string[] sequences2)
        {
            var charSequences1 = new char[sequences1.Length][];
            for (int i = 0; i < sequences1.Length; i++)
            {
                charSequences1[i] = PrepareStringForComparison(sequences1[i]);
            }

            var charSequences2 = new char[sequences2.Length][];
            for (int i = 0; i < sequences2.Length; i++)
            {
                charSequences2[i] = PrepareStringForComparison(sequences2[i]);
            }

            return CalcSimilarityMatrix(charSequences1, charSequences2);
        }
    }

    public static class MatchUtil
    {

        class Sequence
        {
            private readonly List<int> _targetIndexes;
            private readonly List<int> _sourceIndexes;
            private int _score;

            public Sequence(List<int> targetIndexes = null, List<int> sourceIndexes = null, int score = 0)
            {
                _targetIndexes = targetIndexes ?? new List<int>();
                _sourceIndexes = sourceIndexes ?? new List<int>();
                _score = score;
            }

            public int Score
            {
                get => _score;
                set => _score = value;
            }

            public IReadOnlyList<int> TargetIndexes => _targetIndexes;
            public IReadOnlyList<int> SourceIndexes => _sourceIndexes;

            public void AddItem(int targetIndex, int sourceIndex)
            {
                _targetIndexes.Add(targetIndex);
                _sourceIndexes.Add(sourceIndex);
            }

            public void InsertItem(int index, int targetIndex, int sourceIndex)
            {
                _targetIndexes.Insert(index, targetIndex);
                _sourceIndexes.Insert(index, sourceIndex);
            }

            public int GetSourceIndex(int targetIndex)
            {
                int index = _targetIndexes.IndexOf(targetIndex);
                return index >= 0 ? _sourceIndexes[index] : -1;
            }

            public Sequence Clone()
            {
                return new Sequence(new List<int>(_targetIndexes), new List<int>(_sourceIndexes), _score);
            }
        }

        public static List<int> FindAllIndexes<T>(List<T> list, T item, IEqualityComparer<T> comparer = null)
        {
            if (comparer == null)
            {
                comparer = EqualityComparer<T>.Default;
            }

            var indexes = new List<int>();
            for (int i = 0; i < list.Count; i++)
            {
                if (comparer.Equals(list[i], item))
                {
                    indexes.Add(i);
                }
            }

            return indexes;
        }

        private static List<Sequence> SortSequences(List<Sequence> sequences)
        {
            return sequences
                .OrderByDescending(sequence => sequence.TargetIndexes.Count)
                .ThenByDescending(sequence => sequence.Score)
                .ToList();
        }

        private static int CalcSequenceScore(Sequence sequence)
        {
            const int SCORE_MISSED_ITEMS = 1;
            const int SCORE_FOUND_ITEMS = 5;

            int score = 0;
            for (int i = 0; i < sequence.TargetIndexes.Count - 1; i++)
            {
                int targetIndex = sequence.TargetIndexes[i];
                int sourceIndex = sequence.SourceIndexes[i];
                int nextTargetIndex = sequence.TargetIndexes[i + 1];
                int nextSourceIndex = sequence.SourceIndexes[i + 1];

                int targetIndexesDiff = nextTargetIndex - targetIndex;
                int sourceIndexesDiff = nextSourceIndex - sourceIndex;

                score -= (sourceIndexesDiff - 1) * SCORE_MISSED_ITEMS + (targetIndexesDiff - 1) * SCORE_MISSED_ITEMS;
            }
            score += (sequence.TargetIndexes.Count - 1) * SCORE_FOUND_ITEMS;

            return score;
        }

        public static List<int> MatchSequences<T>(List<T> source, List<T> target, IEqualityComparer<T> comparer = null)
        {
            var sequences = new List<Sequence>();

            foreach (var (item, targetIndex) in target.Select((item, index) => (item, index)))
            {
                var sourceItemIndexes = FindAllIndexes(source, item, comparer);

                var sequencesCopy = new List<Sequence>(sequences);

                foreach (var sequence in sequencesCopy)
                {
                    if (sequence.TargetIndexes.Contains(targetIndex))
                    {
                        continue;
                    }

                    var sequenceMaxSourceIndex = sequence.SourceIndexes.Max();

                    var sourceIndex = sourceItemIndexes
                        .Where(index => index > sequenceMaxSourceIndex)
                        .OrderBy(index => index)
                        .DefaultIfEmpty(-1)
                        .FirstOrDefault();

                    if (sourceIndex == -1)
                    {
                        continue;
                    }

                    if (sequenceMaxSourceIndex != 0 && sourceIndex - sequenceMaxSourceIndex > 10)
                    {
                        continue;
                    }

                    var forkedSequence = sequence.Clone();
                    forkedSequence.AddItem(targetIndex, sourceIndex);
                    forkedSequence.Score = CalcSequenceScore(forkedSequence);

                    sequences.Add(forkedSequence);
                }

                sequences = SortSequences(sequences);

                if (sequences.Count > 0 && sequences[0].TargetIndexes.Count > 5)
                {
                    sequences = sequences.Take(5).ToList();
                }
                else
                {
                    foreach (var sourceIndex in sourceItemIndexes)
                    {
                        var sequence = new Sequence();
                        sequence.AddItem(targetIndex, sourceIndex);
                        sequence.Score = CalcSequenceScore(sequence);
                        sequences.Add(sequence);
                    }
                }
            }

            var bestSequence = SortSequences(sequences).FirstOrDefault();

            var sourceIndexes = target.Select((item, targetIndex) => bestSequence?.GetSourceIndex(targetIndex) ?? -1).ToList();

            return sourceIndexes;
        }

    }


    public static class CompareUtil
    {
        public static List<List<double>> CalcParagraphsSimilarityMatrix(List<List<string>> paragraphs1Words, List<List<string>> paragraphs2Words)
        {
            // measure time 
            var start = DateTime.Now;

            var paragraphs1WordsArray = paragraphs1Words.Select(paragraph => paragraph.ToArray()).ToArray();
            var paragraphs2WordsArray = paragraphs2Words.Select(paragraph => paragraph.ToArray()).ToArray();

            var paragraphsSimilarityMatrix = SimilarityUtil.CalcSimilarityMatrix(paragraphs1WordsArray, paragraphs2WordsArray);

            var result = Enumerable.Range(0, paragraphsSimilarityMatrix.GetLength(0))
                .Select(i => Enumerable.Range(0, paragraphsSimilarityMatrix.GetLength(1))
                    .Select(j => paragraphsSimilarityMatrix[i, j])
                    .ToList())
                .ToList();

            // measure time
            var end = DateTime.Now;
            var elapsed = end - start;
            Console.WriteLine("Elapsed time: " + elapsed.TotalMilliseconds);

            return result;
        }

        /*
        * Groups indexes of the similar paragraphs.
        */
        private static List<List<int>> GroupSimilarParagraphs(List<List<string>> paragraphsWords, double similarityThreshold = 0.9)
        {
            var paragraphsIndexesGroups = new List<List<int>>();
            var processedParagraphsIndexes = new HashSet<int>();
            var paragraphsSimilarityMatrix = CalcParagraphsSimilarityMatrix(paragraphsWords, paragraphsWords);

            for (int i = 0; i < paragraphsSimilarityMatrix.Count; i++)
            {
                if (processedParagraphsIndexes.Contains(i))
                {
                    continue;
                }

                var similarParagraphsIndexes = new List<int> { i };
                processedParagraphsIndexes.Add(i);

                for (int j = i + 1; j < paragraphsSimilarityMatrix[i].Count; j++)
                {
                    if (paragraphsSimilarityMatrix[i][j] >= similarityThreshold)
                    {
                        similarParagraphsIndexes.Add(j);
                        processedParagraphsIndexes.Add(j);
                    }
                }

                paragraphsIndexesGroups.Add(similarParagraphsIndexes);
            }

            return paragraphsIndexesGroups;
        }

        /*
        * Calculates IDs for the paragraphs based on the similarity matrix and provided similarity threshold.
        * Finds the highest similarity among all paragraphs and assigns the same ID to the most similar pair.
        * Paragraphs could be considered similar if the similarity is greater than the threshold.
        */
        private static (List<int> SourceIDs, List<int> TargetIDs) BuildParagraphsSimilarityIDs(
            List<List<string>> paragraphs1Words,
            List<List<string>> paragraphs2Words,
            double similarityThreshold = 0.3)
        {
            // Run this in parallel
            var similarParagraphs1GroupsTask = Task.Run(() => GroupSimilarParagraphs(paragraphs1Words));
            var similarParagraphs2GroupsTask = Task.Run(() => GroupSimilarParagraphs(paragraphs2Words));
            var paragraphsSimilarityMatrixTask = Task.Run(() => CalcParagraphsSimilarityMatrix(paragraphs1Words, paragraphs2Words));

            Task.WaitAll(similarParagraphs1GroupsTask, similarParagraphs2GroupsTask, paragraphsSimilarityMatrixTask);

            var similarParagraphs1Groups = similarParagraphs1GroupsTask.Result;
            var similarParagraphs2Groups = similarParagraphs2GroupsTask.Result;
            var paragraphsSimilarityMatrix = paragraphsSimilarityMatrixTask.Result;

            var sourcesCount = paragraphsSimilarityMatrix.Count;
            var targetsCount = paragraphsSimilarityMatrix[0].Count;

            var sourceIDs = Enumerable.Repeat(-1, sourcesCount).ToList();
            var targetIDs = Enumerable.Repeat(-1, targetsCount).ToList();
            var currentID = 100;

            // Pushes the pair to the pairs array if it is unique.
            void PushPairIfUnique(List<(int, int)> pairs, (int, int) pair)
            {
                if (!pairs.Any(p => p.Item1 == pair.Item1 && p.Item2 == pair.Item2))
                {
                    pairs.Add(pair);
                }
            }

            // Function, which will find the highest similarity pair in the matrix and find all other similar pairs.
            List<(int SourceIndex, int TargetIndex)> FindMaxSimilarityIndexes(List<List<double>> matrix)
            {
                double maxSimilarity = 0;
                (int SourceIndex, int TargetIndex)? maxSimilarityPair = null;

                for (int i = 0; i < matrix.Count; i++)
                {
                    for (int j = 0; j < matrix[i].Count; j++)
                    {
                        double similarity = matrix[i][j];
                        if (similarity >= similarityThreshold && similarity > maxSimilarity)
                        {
                            maxSimilarity = similarity;
                            maxSimilarityPair = (SourceIndex: i, TargetIndex: j);
                        }
                    }
                }

                var foundMaxSimilarityIndexes = new List<(int SourceIndex, int TargetIndex)>();

                if (maxSimilarityPair.HasValue)
                {
                    // Find paragraphs1 group, contained index
                    var paragraph1IndexesGroup = similarParagraphs1Groups.FirstOrDefault(group => group.Contains(maxSimilarityPair.Value.SourceIndex)) ?? new List<int>();
                    foreach (var index in paragraph1IndexesGroup)
                    {
                        PushPairIfUnique(foundMaxSimilarityIndexes, (index, maxSimilarityPair.Value.TargetIndex));
                    }

                    // Find paragraphs2 group, contained index
                    var paragraph2IndexesGroup = similarParagraphs2Groups.FirstOrDefault(group => group.Contains(maxSimilarityPair.Value.TargetIndex)) ?? new List<int>();
                    foreach (var index in paragraph2IndexesGroup)
                    {
                        PushPairIfUnique(foundMaxSimilarityIndexes, (maxSimilarityPair.Value.SourceIndex, index));
                    }
                }

                return foundMaxSimilarityIndexes;
            }

            var maxSimilarityIndexes = FindMaxSimilarityIndexes(paragraphsSimilarityMatrix);

            while (maxSimilarityIndexes.Count > 0)
            {
                foreach (var (sourceIndex, targetIndex) in maxSimilarityIndexes)
                {
                    sourceIDs[sourceIndex] = currentID;
                    targetIDs[targetIndex] = currentID;

                    // Put zeroes in the row and column of the matched pair
                    for (int j = 0; j < targetsCount; j++)
                    {
                        paragraphsSimilarityMatrix[sourceIndex][j] = 0;
                    }
                    for (int i = 0; i < sourcesCount; i++)
                    {
                        paragraphsSimilarityMatrix[i][targetIndex] = 0;
                    }
                }

                currentID++;
                maxSimilarityIndexes = FindMaxSimilarityIndexes(paragraphsSimilarityMatrix);
            }

            // Assign unique IDs to any remaining paragraphs without similarities
            void AssignUniqueIDs(List<int> ids)
            {
                for (int i = 0; i < ids.Count; i++)
                {
                    if (ids[i] == -1)
                    {
                        ids[i] = currentID++;
                    }
                }
            }

            AssignUniqueIDs(sourceIDs);
            AssignUniqueIDs(targetIDs);

            return (SourceIDs: sourceIDs, TargetIDs: targetIDs);
        }

        private static (Dictionary<int, int> Matches, Dictionary<int, int> Steady, Dictionary<int, int> Moved) CalcParagraphsMatchesAndMoves(List<int> docParagraphsIDs1, List<int> docParagraphsIDs2)
        {

            // maps source indexes of doc1 (array index) to target indexes of doc2 (array value)
            var docParagraphsIDsMatches = MatchUtil.MatchSequences(docParagraphsIDs2, docParagraphsIDs1);

            var sourceIndexesToMove = new List<int>();
            var targetIndexesInPlace = new List<int>();

            var steadyResults = new Dictionary<int, int>();

            for (int sourceIndex = 0; sourceIndex < docParagraphsIDsMatches.Count; sourceIndex++)
            {
                var targetIndex = docParagraphsIDsMatches[sourceIndex];
                if (targetIndex == -1)
                {
                    sourceIndexesToMove.Add(sourceIndex);
                }
                else
                {
                    targetIndexesInPlace.Add(targetIndex);
                    steadyResults[sourceIndex] = targetIndex;
                }
            }

            var movedResults = new Dictionary<int, int>();

            while (sourceIndexesToMove.Count > 0)
            {
                var sourceIndex = sourceIndexesToMove[0];
                sourceIndexesToMove.RemoveAt(0);
                var paragraphId = docParagraphsIDs1[sourceIndex];

                // find all target indexes for paragraph ID, which are not in place
                var targetIndexes = docParagraphsIDs2
                    .Select((id, index) => new { id, index })
                    .Where(x => x.id == paragraphId && !targetIndexesInPlace.Contains(x.index))
                    .Select(x => x.index)
                    .ToList();

                // get the closest to the source index
                var targetIndex = targetIndexes
                    .OrderBy(index => Math.Abs(sourceIndex - index))
                    .DefaultIfEmpty(-1)
                    .FirstOrDefault();

                if (targetIndex != -1)
                {
                    targetIndexesInPlace.Add(targetIndex);
                    movedResults[sourceIndex] = targetIndex;
                }
            }

            var matches = steadyResults.Concat(movedResults).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            return (Matches: matches, Steady: steadyResults, Moved: movedResults);
        }

        public static (Dictionary<int, int> Matches, Dictionary<int, int> Steady, Dictionary<int, int> Moved) CompareDocumentsParagraphs(List<List<string>> paragraphs1Words, List<List<string>> paragraphs2Words, double similarityThreshold = 0.3)
        {
            var (docParagraphsIDs1, docParagraphsIDs2) = BuildParagraphsSimilarityIDs(paragraphs1Words, paragraphs2Words, similarityThreshold);

            return CalcParagraphsMatchesAndMoves(docParagraphsIDs1, docParagraphsIDs2);
        }


    }

}