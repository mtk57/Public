using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using Tesseract;

namespace SimpleOCRforBase32
{
    public partial class Form1 : Form
    {
        public Form1 ()
        {
            InitializeComponent();
            InitializeEvents();
        }

        private static void AddCandidateFromRawText ( List<CandidateResult> candidates, string rawText, float confidence, bool numericMode, string source )
        {
            if ( candidates == null )
            {
                return;
            }

            var normalized = NormalizeText( rawText ?? string.Empty );
            if ( string.IsNullOrWhiteSpace( normalized ) )
            {
                return;
            }

            candidates.Add( new CandidateResult( normalized, confidence, numericMode, source ) );
        }

        private static void AddCandidateFromNormalizedLines ( List<CandidateResult> candidates, IEnumerable<string> lines, float confidence, bool numericMode, string source )
        {
            if ( candidates == null || lines == null )
            {
                return;
            }

            var materialized = lines
                .Where( line => !string.IsNullOrWhiteSpace( line ) )
                .ToList();

            if ( materialized.Count == 0 )
            {
                return;
            }

            var text = string.Join( Environment.NewLine, materialized );
            candidates.Add( new CandidateResult( text, confidence, numericMode, source ) );
        }

        private static string MergeCandidates ( List<CandidateResult> candidates )
        {
            if ( candidates == null || candidates.Count == 0 )
            {
                return string.Empty;
            }

            var activeCandidates = candidates
                .Where( c => c.LineInfos.Length > 0 )
                .ToList();

            if ( activeCandidates.Count == 0 )
            {
                activeCandidates = candidates;
            }

            var lineCount = activeCandidates.Max( c => c.LineInfos.Length );
            var mergedLines = new List<string>( lineCount );

            for ( var lineIndex = 0; lineIndex < lineCount; lineIndex++ )
            {
                var lineCandidates = activeCandidates
                    .Where( c => lineIndex < c.LineInfos.Length && c.LineInfos[lineIndex].IsChunkComplete )
                    .ToList();

                if ( lineCandidates.Count == 0 )
                {
                    continue;
                }

                var directMatch = lineCandidates
                    .Select( c => ( Candidate: c, Info: c.GetLineInfo( lineIndex ) ) )
                    .Where( x => x.Info != null && x.Info.ChecksumMatches )
                    .OrderByDescending( x => x.Candidate.Weight )
                    .FirstOrDefault();

                if ( directMatch.Info != null )
                {
                    var formattedDirect = $"{FormatPrefix( directMatch.Info.Chunk )}{ChecksumSeparator}{directMatch.Info.Checksum}";
                    mergedLines.Add( formattedDirect );
                    continue;
                }

                var sanitizedLength = lineCandidates
                    .Select( c => c.LineInfos[lineIndex].CombinedLength )
                    .Where( length => length >= ChunkLength )
                    .DefaultIfEmpty( 0 )
                    .Max();

                if ( sanitizedLength < ChunkLength )
                {
                    continue;
                }

                sanitizedLength = Math.Min( sanitizedLength, ChunkLength + 2 );

                var votesByPosition = new Dictionary<int, List<CharVote>>();

                foreach ( var candidate in lineCandidates )
                {
                    var info = candidate.LineInfos[lineIndex];
                    var combined = info.Combined;
                    if ( string.IsNullOrEmpty( combined ) )
                    {
                        continue;
                    }

                    var length = Math.Min( sanitizedLength, combined.Length );
                    for ( var posLoop = 0; posLoop < length; posLoop++ )
                    {
                        var chLoop = combined[posLoop];
                        if ( !char.IsLetterOrDigit( chLoop ) )
                        {
                            continue;
                        }

                        if ( !votesByPosition.TryGetValue( posLoop, out var list ) )
                        {
                            list = new List<CharVote>();
                            votesByPosition[posLoop] = list;
                        }

                        var weightMultiplier = info.ChecksumMatches ? 1.1f : 1.0f;
                        list.Add( new CharVote( chLoop, candidate.Weight * weightMultiplier, candidate ) );
                    }
                }

                if ( votesByPosition.Count == 0 )
                {
                    continue;
                }

                var chunkLength = Math.Min( ChunkLength, sanitizedLength );
                var chunkChars = new char[chunkLength];

                var orderedLineCandidates = lineCandidates
                    .OrderByDescending( c => c.Weight )
                    .ToList();

                for ( var posLoop = 0; posLoop < chunkLength; posLoop++ )
                {
                    if ( !votesByPosition.TryGetValue( posLoop, out var votes ) || votes.Count == 0 )
                    {
                        chunkChars[posLoop] = GetFallbackCharacter( orderedLineCandidates, lineIndex, posLoop );
                        continue;
                    }

                    chunkChars[posLoop] = SelectCharacterForPosition( votes );
                }

                ApplyBigramAdjustments( chunkChars, votesByPosition );

                var targetChecksum = DetermineTargetChecksum( lineCandidates, lineIndex );
                var chunk = new string( chunkChars );
                var checksum = ComputeChecksum( chunk );

                if ( !string.IsNullOrEmpty( targetChecksum )
                    && targetChecksum.Length == 2
                    && !string.Equals( checksum, targetChecksum, StringComparison.OrdinalIgnoreCase )
                    && TryAdjustChunkWithChecksum( chunkChars, votesByPosition, targetChecksum, out var correctedChunk ) )
                {
                    chunk = correctedChunk;
                    chunkChars = correctedChunk.ToCharArray();
                    checksum = targetChecksum;
                }
                else
                {
                    checksum = ComputeChecksum( chunk );
                }

                var bestAligned = lineCandidates
                    .Select( c => ( Candidate: c, Info: c.GetLineInfo( lineIndex ) ) )
                    .Where( x => x.Info != null )
                    .Select( x => new
                    {
                        x.Info,
                        Weight = x.Candidate.Weight,
                        Distance = ComputeHammingDistance( chunk, x.Info.Chunk )
                    } )
                    .OrderBy( x => x.Distance )
                    .ThenByDescending( x => x.Weight )
                    .FirstOrDefault();

                if ( bestAligned?.Info != null && bestAligned.Distance > 0 && bestAligned.Distance <= 4 )
                {
                    var candidateChecksum = ComputeChecksum( bestAligned.Info.Chunk );
                    var matchesTarget = string.IsNullOrEmpty( targetChecksum )
                        ? bestAligned.Info.ChecksumMatches
                        : string.Equals( candidateChecksum, targetChecksum, StringComparison.OrdinalIgnoreCase );

                    if ( matchesTarget )
                    {
                        chunk = bestAligned.Info.Chunk;
                        checksum = candidateChecksum;
                    }
                }

                var formatted = $"{FormatPrefix( chunk )}{ChecksumSeparator}{checksum}";
                mergedLines.Add( formatted );
            }

            if ( mergedLines.Count == 0 )
            {
                var fallback = candidates
                    .OrderByDescending( c => c.ValidLineCount )
                    .ThenByDescending( c => c.Confidence )
                    .FirstOrDefault();
                return fallback?.Text ?? string.Empty;
            }

            return string.Join( Environment.NewLine, mergedLines );
        }


        private static char SelectCharacterForPosition ( List<CharVote> votes )
        {
            if ( votes == null || votes.Count == 0 )
            {
                return ' ';
            }

            var summaries = votes
                .GroupBy( v => v.Character )
                .Select( g => new VoteSummary( g.Key, g.ToList() ) )
                .OrderByDescending( summary => summary.Weight )
                .ToList();

            var primary = summaries[0];
            if ( summaries.Count == 1 )
            {
                return primary.Character;
            }

            var secondary = summaries[1];
            var diff = primary.Weight - secondary.Weight;

            if ( IsAmbiguousPair( primary.Character, secondary.Character ) && diff < 0.05f )
            {
                return ResolveAmbiguousPair( primary, secondary );
            }

            return primary.Character;
        }

        private static char GetFallbackCharacter ( IReadOnlyList<CandidateResult> orderedCandidates, int lineIndex, int position )
        {
            if ( orderedCandidates == null )
            {
                return ' ';
            }

            foreach ( var candidate in orderedCandidates )
            {
                var info = candidate.GetLineInfo( lineIndex );
                if ( info == null )
                {
                    continue;
                }

                if ( position < info.Chunk.Length )
                {
                    return info.Chunk[position];
                }
            }

            return ' ';
        }

        private static void ApplyBigramAdjustments ( char[] chunkChars, Dictionary<int, List<CharVote>> votesByPosition )
        {
            if ( chunkChars == null || chunkChars.Length < 2 )
            {
                return;
            }

            for ( var i = 0; i < chunkChars.Length - 1; i++ )
            {
                var current = chunkChars[i];
                var next = chunkChars[i + 1];

                if ( ( current == '5' && next == 'S' ) || ( current == 'S' && next == '5' ) )
                {
                    var weight5 = GetCharWeight( votesByPosition, i, '5' ) + GetCharWeight( votesByPosition, i + 1, '5' );
                    var weightS = GetCharWeight( votesByPosition, i, 'S' ) + GetCharWeight( votesByPosition, i + 1, 'S' );

                    if ( weight5 > weightS * 1.1f )
                    {
                        chunkChars[i] = '5';
                        chunkChars[i + 1] = '5';
                    }
                    else if ( weightS > weight5 * 1.1f )
                    {
                        chunkChars[i] = 'S';
                        chunkChars[i + 1] = 'S';
                    }
                }
            }
        }

        private static float GetCharWeight ( Dictionary<int, List<CharVote>> votesByPosition, int position, char character )
        {
            if ( votesByPosition == null || !votesByPosition.TryGetValue( position, out var votes ) )
            {
                return 0f;
            }

            return votes
                .Where( vote => vote.Character == character )
                .Sum( vote => vote.Weight );
        }

        private static int ComputeHammingDistance ( string left, string right )
        {
            if ( string.IsNullOrEmpty( left ) || string.IsNullOrEmpty( right ) )
            {
                return int.MaxValue;
            }

            var length = Math.Min( left.Length, right.Length );
            var distance = 0;
            for ( var i = 0; i < length; i++ )
            {
                if ( left[i] != right[i] )
                {
                    distance++;
                }
            }

            distance += Math.Abs( left.Length - right.Length );
            return distance;
        }

        private static string DetermineTargetChecksum ( IEnumerable<CandidateResult> lineCandidates, int lineIndex )
        {
            if ( lineCandidates == null )
            {
                return null;
            }

            var votes = new Dictionary<string, float>( StringComparer.Ordinal );

            foreach ( var candidate in lineCandidates )
            {
                var info = candidate.GetLineInfo( lineIndex );
                if ( info == null || !info.HasChecksum )
                {
                    continue;
                }

                if ( !votes.TryGetValue( info.Checksum, out var weight ) )
                {
                    votes[info.Checksum] = 0f;
                }

                var weightContribution = candidate.Weight * ( info.ChecksumMatches ? 1.2f : 1.0f );
                votes[info.Checksum] += weightContribution;
            }

            if ( votes.Count == 0 )
            {
                return null;
            }

            return votes
                .OrderByDescending( kvp => kvp.Value )
                .First()
                .Key;
        }

        private static bool TryAdjustChunkWithChecksum ( char[] chunkChars, Dictionary<int, List<CharVote>> votesByPosition, string targetChecksum, out string corrected )
        {
            corrected = null;

            if ( chunkChars == null || chunkChars.Length == 0 || string.IsNullOrEmpty( targetChecksum ) )
            {
                return false;
            }

            var currentChecksum = ComputeChecksum( new string( chunkChars ) );
            if ( string.Equals( currentChecksum, targetChecksum, StringComparison.OrdinalIgnoreCase ) )
            {
                corrected = new string( chunkChars );
                return true;
            }

            var positions = new List<AmbiguousPosition>();

            for ( var pos = 0; pos < chunkChars.Length; pos++ )
            {
                var alternatives = GetAmbiguousAlternatives( pos, chunkChars[pos], votesByPosition );
                if ( alternatives.Count > 0 )
                {
                    positions.Add( new AmbiguousPosition( pos, alternatives ) );
                }
            }

            if ( positions.Count == 0 )
            {
                return false;
            }

            positions = positions
                .OrderByDescending( p => p.Alternatives.Count > 0 ? p.Alternatives[0].Weight : 0f )
                .ToList();

            var maxDepth = Math.Min( 3, positions.Count );
            if ( TryAdjustRecursive( chunkChars, positions, 0, maxDepth, targetChecksum, out corrected ) )
            {
                return true;
            }

            corrected = null;
            return false;
        }

        private static List<AlternativeChar> GetAmbiguousAlternatives ( int position, char currentChar, Dictionary<int, List<CharVote>> votesByPosition )
        {
            var result = new List<AlternativeChar>();

            List<CharVote> votes = null;
            if ( votesByPosition != null )
            {
                votesByPosition.TryGetValue( position, out votes );
            }

            if ( votes != null )
            {
                result.AddRange( votes
                    .GroupBy( vote => vote.Character )
                    .Select( g => new AlternativeChar( g.Key, g.Sum( vote => vote.Weight ) ) )
                    .Where( alt => IsAmbiguousPair( alt.Character, currentChar ) && alt.Character != currentChar )
                    .OrderByDescending( alt => alt.Weight )
                    .Take( 2 ) );
            }

            if ( AmbiguityMap.TryGetValue( currentChar, out var mappedChars ) )
            {
                var baseWeight = votes?.Where( vote => vote.Character == currentChar ).Sum( vote => vote.Weight ) ?? 1f;
                foreach ( var mapped in mappedChars )
                {
                    if ( result.Any( alt => alt.Character == mapped ) )
                    {
                        continue;
                    }

                    var mappedWeight = baseWeight * 0.8f;
                    if ( votes != null )
                    {
                        var existing = votes
                            .Where( vote => vote.Character == mapped )
                            .Sum( vote => vote.Weight );
                        if ( existing > 0f )
                        {
                            mappedWeight = Math.Max( mappedWeight, existing );
                        }
                    }

                    result.Add( new AlternativeChar( mapped, mappedWeight ) );
                }
            }

            return result;
        }

        private static bool TryAdjustRecursive ( char[] chunkChars, List<AmbiguousPosition> positions, int index, int depthRemaining, string targetChecksum, out string corrected )
        {
            var currentChecksum = ComputeChecksum( new string( chunkChars ) );
            if ( string.Equals( currentChecksum, targetChecksum, StringComparison.OrdinalIgnoreCase ) )
            {
                corrected = new string( chunkChars );
                return true;
            }

            if ( depthRemaining == 0 || index >= positions.Count )
            {
                corrected = null;
                return false;
            }

            for ( var i = index; i < positions.Count; i++ )
            {
                var position = positions[i];
                var original = chunkChars[position.Index];

                foreach ( var alternative in position.Alternatives )
                {
                    if ( alternative.Character == original )
                    {
                        continue;
                    }

                    chunkChars[position.Index] = alternative.Character;
                    if ( TryAdjustRecursive( chunkChars, positions, i + 1, depthRemaining - 1, targetChecksum, out corrected ) )
                    {
                        return true;
                    }
                }

                chunkChars[position.Index] = original;
            }

            corrected = null;
            return false;
        }


        private static bool IsAmbiguousPair ( char a, char b )
        {
            return AmbiguityMap.TryGetValue( a, out var mapped ) && mapped.Contains( b );
        }

        private static char ResolveAmbiguousPair ( VoteSummary primary, VoteSummary secondary )
        {
            if ( primary == null || secondary == null )
            {
                return primary?.Character ?? secondary?.Character ?? ' ';
            }

            if ( IsDigit( primary.Character ) != IsDigit( secondary.Character ) )
            {
                var digit = IsDigit( primary.Character ) ? primary : secondary;
                var letter = IsDigit( primary.Character ) ? secondary : primary;

                if ( digit.NumericWeight > letter.AlphaWeight * 1.1f )
                {
                    return digit.Character;
                }

                if ( letter.AlphaWeight > digit.NumericWeight * 1.05f )
                {
                    return letter.Character;
                }
            }

            if ( IsIOrJPair( primary.Character, secondary.Character ) )
            {
                if ( primary.AlphaWeight > secondary.AlphaWeight * 1.05f )
                {
                    return primary.Character;
                }

                if ( secondary.AlphaWeight > primary.AlphaWeight * 1.05f )
                {
                    return secondary.Character;
                }
            }

            return primary.Weight >= secondary.Weight
                ? primary.Character
                : secondary.Character;
        }

        private static bool IsDigit ( char c ) =>
            c >= '0' && c <= '9';

        private static bool IsIOrJPair ( char a, char b ) =>
            ( a == 'I' && b == 'J' ) || ( a == 'J' && b == 'I' );

        private void InitializeEvents ()
        {
            txtImgFilePath.AllowDrop = true;
            txtImgFilePath.DragEnter += TxtImgFilePath_DragEnter;
            txtImgFilePath.DragDrop += TxtImgFilePath_DragDrop;
            btnRefImgFilePath.Click += BtnRefImgFilePath_Click;
            btnStart.Click += BtnStart_Click;
        }

        private void TxtImgFilePath_DragEnter ( object sender, DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
                e.Effect = files != null && files.Length > 0 && File.Exists( files[0] )
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TxtImgFilePath_DragDrop ( object sender, DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
                if ( files != null && files.Length > 0 && File.Exists( files[0] ) )
                {
                    txtImgFilePath.Text = files[0];
                }
            }
        }

        private void BtnRefImgFilePath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new OpenFileDialog() )
            {
                dialog.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff|All Files|*.*";
                dialog.Title = "画像ファイルを選択してください";
                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtImgFilePath.Text = dialog.FileName;
                }
            }
        }

        private async void BtnStart_Click ( object sender, EventArgs e )
        {
            var filePath = txtImgFilePath.Text;
            if ( string.IsNullOrWhiteSpace( filePath ) || !File.Exists( filePath ) )
            {
                MessageBox.Show( this, "有効な画像ファイルパスを指定してください。", "ファイル未指定", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            btnStart.Enabled = false;
            txtResultOCR.Clear();

            try
            {
                var ocrText = await Task.Run( () => RunOcr( filePath ) );
                txtResultOCR.Text = ocrText;
            }
            catch ( DirectoryNotFoundException ex )
            {
                MessageBox.Show( this, ex.Message, "tessdata が見つかりません", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"OCR処理でエラーが発生しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
            finally
            {
                btnStart.Enabled = true;
            }
        }

        private string RunOcr ( string filePath )
        {
            var tessdataPath = Path.Combine( AppDomain.CurrentDomain.BaseDirectory, "tessdata" );
            if ( !Directory.Exists( tessdataPath ) )
            {
                throw new DirectoryNotFoundException( $"tessdata フォルダーが見つかりません。以下に配置してください。\n{tessdataPath}" );
            }

            PreprocessResult preprocessResult = null;

            try
            {
                preprocessResult = PreprocessImage( filePath );
            }
            catch
            {
                preprocessResult = null;
            }

            try
            {
                using ( var engine = new TesseractEngine( tessdataPath, "eng", EngineMode.Default ) )
                {
                    engine.SetVariable( "tessedit_char_whitelist", AllowedCharacters );
                    engine.SetVariable( "preserve_interword_spaces", "1" );
                    engine.SetVariable( "load_system_dawg", "0" );
                    engine.SetVariable( "load_freq_dawg", "0" );
                    engine.SetVariable( "wordrec_enable_assoc", "0" );
                    engine.SetVariable( "language_model_penalty_non_dict_word", "0.15" );
                    engine.SetVariable( "language_model_penalty_non_freq_dict_word", "0.25" );
                    engine.DefaultPageSegMode = PageSegMode.SingleColumn;

                    var candidates = new List<CandidateResult>();

                    var sources = new List<ProcessingSource>();

                    if ( preprocessResult != null
                        && !string.IsNullOrEmpty( preprocessResult.ProcessedImagePath )
                        && File.Exists( preprocessResult.ProcessedImagePath ) )
                    {
                        sources.Add( new ProcessingSource(
                            preprocessResult.ProcessedImagePath,
                            preprocessResult.LineRectangles,
                            "preprocessed" ) );
                    }

                    sources.Add( new ProcessingSource(
                        filePath,
                        Array.Empty<OpenCvSharp.Rect>(),
                        "original" ) );

                    foreach ( var numericMode in new[] { true, false } )
                    {
                        engine.SetVariable( "classify_bln_numeric_mode", numericMode ? "1" : "0" );

                        foreach ( var source in sources )
                        {
                            using ( var pix = Pix.LoadFromFile( source.Path ) )
                            {
                                using ( var page = engine.Process( pix, PageSegMode.SingleColumn ) )
                                {
                                    AddCandidateFromRawText(
                                        candidates,
                                        page.GetText(),
                                        page.GetMeanConfidence(),
                                        numericMode,
                                        $"{source.Name}:column" );
                                }

                                if ( source.LineRectangles.Count > 0 )
                                {
                                    var lineTexts = new List<string>();
                                    var lineConfidences = new List<float>();

                                    foreach ( var rect in source.LineRectangles )
                                    {
                                        using ( var page = engine.Process( pix, new Tesseract.Rect( rect.X, rect.Y, rect.Width, rect.Height ), PageSegMode.SingleLine ) )
                                        {
                                            var normalizedLine = NormalizeLine( page.GetText() ?? string.Empty );
                                            if ( !string.IsNullOrWhiteSpace( normalizedLine ) )
                                            {
                                                lineTexts.Add( normalizedLine );
                                                lineConfidences.Add( page.GetMeanConfidence() );
                                            }
                                        }
                                    }

                                    if ( lineTexts.Count > 0 )
                                    {
                                        var averageConfidence = lineConfidences.Count > 0
                                            ? (float)lineConfidences.Average()
                                            : 0f;
                                        AddCandidateFromNormalizedLines(
                                            candidates,
                                            lineTexts,
                                            averageConfidence,
                                            numericMode,
                                            $"{source.Name}:lines" );
                                    }
                                }
                            }
                        }
                    }

                    var mergedText = MergeCandidates( candidates );
                    return mergedText;
                }
            }
            finally
            {
                if ( preprocessResult != null
                    && !string.IsNullOrEmpty( preprocessResult.ProcessedImagePath )
                    && File.Exists( preprocessResult.ProcessedImagePath ) )
                {
                    File.Delete( preprocessResult.ProcessedImagePath );
                }
            }
        }

        private const int ChunkLength = 32;
        private const char ChecksumSeparator = '-';
        private const string ChecksumAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";

        private static readonly Dictionary<char, char[]> AmbiguityMap = new Dictionary<char, char[]>
        {
            { '5', new[] { 'S' } },
            { 'S', new[] { '5' } },
            { '4', new[] { 'A' } },
            { 'A', new[] { '4' } },
            { 'I', new[] { 'J', '1', 'L' } },
            { 'J', new[] { 'I', 'L' } },
            { '1', new[] { 'I', 'L' } },
            { 'L', new[] { '1', 'I' } },
            { '2', new[] { 'Z' } },
            { 'Z', new[] { '2' } }
        };

        private static readonly string AllowedCharacters = $"{ChecksumAlphabet}{ChecksumSeparator}";
        private static readonly HashSet<char> AllowedCharacterSet = new HashSet<char>( AllowedCharacters );
        private static readonly Dictionary<char, char> SoftEquivalents = new Dictionary<char, char>
        {
            { '0', 'O' },
            { '1', 'I' },
            { '8', 'B' }
        };

        private static string NormalizeText ( string rawText )
        {
            if ( string.IsNullOrWhiteSpace( rawText ) )
            {
                return string.Empty;
            }

            var lines = rawText
                .Split( new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries )
                .Select( NormalizeLine )
                .Where( line => !string.IsNullOrWhiteSpace( line ) )
                .ToArray();

            return string.Join( Environment.NewLine, lines );
        }

        private static string NormalizeLine ( string rawLine )
        {
            if ( string.IsNullOrWhiteSpace( rawLine ) )
            {
                return string.Empty;
            }

            var sanitized = SanitizeLine( rawLine );
            if ( sanitized.Length < ChunkLength )
            {
                return string.Empty;
            }

            if ( sanitized.Length > ChunkLength + 2 )
            {
                sanitized = sanitized.Substring( 0, ChunkLength + 2 );
            }

            return sanitized;
        }

        private static string FormatPrefix ( string prefix )
        {
            if ( string.IsNullOrEmpty( prefix ) )
            {
                return string.Empty;
            }

            var groups = new List<string>();
            for ( var i = 0; i < prefix.Length; i += 4 )
            {
                var length = Math.Min( 4, prefix.Length - i );
                groups.Add( prefix.Substring( i, length ) );
            }

            return string.Join( " ", groups );
        }

        private static string ComputeChecksum ( string chunk )
        {
            if ( string.IsNullOrEmpty( chunk ) )
            {
                return new string( ChecksumAlphabet[0], 2 );
            }

            unchecked
            {
                var hash = 0;
                foreach ( var ch in chunk )
                {
                    hash = ( hash * 33 ) ^ ch;
                }

                hash &= 0x3FF; // 10 bits
                var first = hash / ChecksumAlphabet.Length;
                var second = hash % ChecksumAlphabet.Length;
                return new string( new[] { ChecksumAlphabet[first], ChecksumAlphabet[second] } );
            }
        }

        private static string SanitizeLine ( string formattedLine )
        {
            if ( string.IsNullOrWhiteSpace( formattedLine ) )
            {
                return string.Empty;
            }

            var builder = new System.Text.StringBuilder( formattedLine.Length );

            foreach ( var ch in formattedLine )
            {
                var upper = char.ToUpperInvariant( ch );
                if ( upper == ChecksumSeparator )
                {
                    continue;
                }

                if ( SoftEquivalents.TryGetValue( upper, out var mapped ) )
                {
                    builder.Append( mapped );
                    continue;
                }

                if ( AllowedCharacterSet.Contains( upper ) && char.IsLetterOrDigit( upper ) )
                {
                    builder.Append( upper );
                }
            }

            return builder.ToString();
        }

        private static PreprocessResult PreprocessImage ( string originalPath )
        {
            var tempPath = Path.Combine( Path.GetTempPath(), $"simpleocr-pre-{Guid.NewGuid():N}.png" );
            var lineRects = new List<OpenCvSharp.Rect>();

            using ( var src = Cv2.ImRead( originalPath, ImreadModes.Color ) )
            {
                if ( src.Empty() )
                {
                    throw new InvalidOperationException( "画像を読み込めませんでした。" );
                }

                using ( var gray = new Mat() )
                using ( var blurred = new Mat() )
                using ( var binary = new Mat() )
                using ( var closed = new Mat() )
                using ( var scaled = new Mat() )
                {
                    Cv2.CvtColor( src, gray, ColorConversionCodes.BGR2GRAY );
                    Cv2.GaussianBlur( gray, blurred, new OpenCvSharp.Size( 3, 3 ), 0 );
                    Cv2.AdaptiveThreshold(
                        blurred,
                        binary,
                        maxValue: 255,
                        adaptiveMethod: AdaptiveThresholdTypes.MeanC,
                        thresholdType: ThresholdTypes.Binary,
                        blockSize: 17,
                        c: 10 );

                    using ( var kernel = Cv2.GetStructuringElement( MorphShapes.Rect, new OpenCvSharp.Size( 2, 2 ) ) )
                    {
                        Cv2.MorphologyEx( binary, closed, MorphTypes.Close, kernel );
                    }

                    Cv2.Resize( closed, scaled, new OpenCvSharp.Size(), 1.6, 1.6, InterpolationFlags.Linear );
                    var width = scaled.Cols;
                    var height = scaled.Rows;
                    var rowThreshold = Math.Max( 1, width / 25 );
                    var minLineHeight = Math.Max( 5, height / 200 );

                    var y = 0;
                    while ( y < height )
                    {
                        int rowCount;
                        using ( var row = scaled.Row( y ) )
                        {
                            rowCount = Cv2.CountNonZero( row );
                        }

                        if ( rowCount > rowThreshold )
                        {
                            var start = y;
                            do
                            {
                                y++;
                                if ( y >= height )
                                {
                                    break;
                                }

                                using ( var nextRow = scaled.Row( y ) )
                                {
                                    rowCount = Cv2.CountNonZero( nextRow );
                                }
                            }
                            while ( rowCount > rowThreshold );

                            var end = Math.Min( height - 1, y - 1 );
                            var top = Math.Max( 0, start - 2 );
                            var bottom = Math.Min( height - 1, end + 2 );
                            var lineHeight = bottom - top + 1;
                            if ( lineHeight < minLineHeight )
                            {
                                continue;
                            }

                            var roiRect = new OpenCvSharp.Rect( 0, top, width, lineHeight );
                            using ( var roi = new Mat( scaled, roiRect ) )
                            {
                                var columnCounts = new int[width];
                                for ( var x = 0; x < width; x++ )
                                {
                                    using ( var col = roi.Col( x ) )
                                    {
                                        columnCounts[x] = Cv2.CountNonZero( col );
                                    }
                                }

                                var columnThreshold = Math.Max( 1, lineHeight / 8 );
                                var left = 0;
                                var right = width - 1;

                                while ( left < width && columnCounts[left] <= columnThreshold )
                                {
                                    left++;
                                }

                                while ( right >= 0 && columnCounts[right] <= columnThreshold )
                                {
                                    right--;
                                }

                                if ( right <= left )
                                {
                                    lineRects.Add( roiRect );
                                }
                                else
                                {
                                    const int margin = 2;
                                    var adjustedLeft = Math.Max( 0, left - margin );
                                    var adjustedRight = Math.Min( width - 1, right + margin );
                                    var rect = new OpenCvSharp.Rect(
                                        adjustedLeft,
                                        top,
                                        adjustedRight - adjustedLeft + 1,
                                        lineHeight );
                                    lineRects.Add( rect );
                                }
                            }
                        }
                        else
                        {
                            y++;
                        }
                    }

                    Cv2.ImWrite( tempPath, scaled );
                }
            }

            lineRects.Sort( ( a, b ) => a.Y.CompareTo( b.Y ) );
            return new PreprocessResult( tempPath, lineRects );
        }

        private sealed class ProcessingSource
        {
            public ProcessingSource ( string path, IReadOnlyList<OpenCvSharp.Rect> lineRectangles, string name )
            {
                Path = path;
                LineRectangles = lineRectangles ?? Array.Empty<OpenCvSharp.Rect>();
                Name = name ?? string.Empty;
            }

            public string Path { get; }
            public IReadOnlyList<OpenCvSharp.Rect> LineRectangles { get; }
            public string Name { get; }
        }

        private sealed class CandidateResult
        {
            public CandidateResult ( string text, float confidence, bool numericMode, string source )
            {
                Text = text ?? string.Empty;
                Confidence = Math.Max( confidence, 0f );
                NumericMode = numericMode;
                Source = source ?? string.Empty;

                Lines = Text
                    .Split( new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries )
                    .ToArray();

                LineInfos = Lines
                    .Select( CreateLineInfo )
                    .Where( info => info != null )
                    .ToArray();

                ValidLineCount = LineInfos.Count( info => info.IsChunkComplete );
                ChecksumMatchLineCount = LineInfos.Count( info => info.ChecksumMatches );
                Weight = Confidence
                    + ValidLineCount * 0.001f
                    + ChecksumMatchLineCount * 0.01f;
            }

            public string Text { get; }
            public float Confidence { get; }
            public bool NumericMode { get; }
            public string Source { get; }
            public string[] Lines { get; }
            public CandidateLineInfo[] LineInfos { get; }
            public int ValidLineCount { get; }
            public int ChecksumMatchLineCount { get; }
            public float Weight { get; }

            public CandidateLineInfo GetLineInfo ( int index )
            {
                return index >= 0 && index < LineInfos.Length
                    ? LineInfos[index]
                    : null;
            }

            private static CandidateLineInfo CreateLineInfo ( string line )
            {
                var sanitized = SanitizeLine( line );
                if ( string.IsNullOrEmpty( sanitized ) )
                {
                    return null;
                }

                if ( sanitized.Length > ChunkLength + 2 )
                {
                    sanitized = sanitized.Substring( 0, ChunkLength + 2 );
                }

                if ( sanitized.Length < ChunkLength )
                {
                    return null;
                }

                var chunk = sanitized.Substring( 0, ChunkLength );
                var checksum = sanitized.Length >= ChunkLength + 2
                    ? sanitized.Substring( ChunkLength, 2 )
                    : string.Empty;

                var checksumMatches = checksum.Length == 2
                    && string.Equals( ComputeChecksum( chunk ), checksum, StringComparison.OrdinalIgnoreCase );

                return new CandidateLineInfo( chunk, checksum, checksumMatches );
            }
        }

        private sealed class CandidateLineInfo
        {
            public CandidateLineInfo ( string chunk, string checksum, bool checksumMatches )
            {
                Chunk = chunk ?? string.Empty;
                Checksum = checksum ?? string.Empty;
                ChecksumMatches = checksumMatches;
                Combined = Chunk + Checksum;
            }

            public string Chunk { get; }
            public string Checksum { get; }
            public bool ChecksumMatches { get; }
            public string Combined { get; }
            public int CombinedLength => Combined.Length;
            public bool HasChecksum => Checksum.Length == 2;
            public bool IsChunkComplete => Chunk.Length >= ChunkLength;
        }

        private sealed class CharVote
        {
            public CharVote ( char character, float weight, CandidateResult candidate )
            {
                Character = character;
                Weight = weight;
                Candidate = candidate;
            }

            public char Character { get; }
            public float Weight { get; }
            public CandidateResult Candidate { get; }
        }

        private sealed class VoteSummary
        {
            public VoteSummary ( char character, List<CharVote> votes )
            {
                Character = character;
                Votes = votes ?? new List<CharVote>();
                Weight = Votes.Sum( v => v.Weight );
                NumericWeight = Votes.Where( v => v.Candidate.NumericMode ).Sum( v => v.Weight );
                AlphaWeight = Votes.Where( v => !v.Candidate.NumericMode ).Sum( v => v.Weight );
            }

            public char Character { get; }
            public List<CharVote> Votes { get; }
            public float Weight { get; }
            public float NumericWeight { get; }
            public float AlphaWeight { get; }
        }

        private sealed class AmbiguousPosition
        {
            public AmbiguousPosition ( int index, IReadOnlyList<AlternativeChar> alternatives )
            {
                Index = index;
                Alternatives = alternatives?.ToList() ?? new List<AlternativeChar>();
            }

            public int Index { get; }
            public List<AlternativeChar> Alternatives { get; }
        }

        private sealed class AlternativeChar
        {
            public AlternativeChar ( char character, float weight )
            {
                Character = character;
                Weight = weight;
            }

            public char Character { get; }
            public float Weight { get; }
        }

        private sealed class PreprocessResult
        {
            public PreprocessResult ( string processedImagePath, IReadOnlyList<OpenCvSharp.Rect> lineRectangles )
            {
                ProcessedImagePath = processedImagePath;
                LineRectangles = lineRectangles ?? Array.Empty<OpenCvSharp.Rect>();
            }

            public string ProcessedImagePath { get; }
            public IReadOnlyList<OpenCvSharp.Rect> LineRectangles { get; }
        }
    }
}
