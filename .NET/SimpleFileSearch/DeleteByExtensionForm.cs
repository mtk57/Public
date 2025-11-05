using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SimpleFileSearch
{
    public partial class DeleteByExtensionForm : Form
    {
        private const string NoExtensionDisplayName = "(なし)";
        private readonly Dictionary<string, List<SearchResult>> _extensionGroups;
        private readonly List<string> _deletedFiles = new List<string>();
        private readonly HashSet<string> _deletedFileSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private List<ExtensionGroup> _filteredGroups = new List<ExtensionGroup>();

        public DeleteByExtensionForm ( IEnumerable<SearchResult> results )
        {
            if ( results == null )
            {
                throw new ArgumentNullException( nameof( results ) );
            }

            InitializeComponent();

            clmCount.ValueType = typeof(int);
            clmCount.DefaultCellStyle.Format = "N0";
            clmCount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            _extensionGroups = results
                .GroupBy( r => r.Extension ?? string.Empty, StringComparer.OrdinalIgnoreCase )
                .ToDictionary( g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase );

            txtExtFilter.TextChanged += TxtExtFilter_TextChanged;
            dataGridViewResults.SelectionChanged += DataGridViewResults_SelectionChanged;
            dataGridViewResults.KeyDown += DataGridViewResults_KeyDown;
            btnDeleteSelected.Click += btnDeleteSelected_Click;

            ApplyFilter();
        }

        public IReadOnlyList<string> DeletedFiles => _deletedFiles;

        private void TxtExtFilter_TextChanged ( object sender, EventArgs e )
        {
            ApplyFilter();
        }

        private void DataGridViewResults_SelectionChanged ( object sender, EventArgs e )
        {
            UpdateDeleteButtonState();
        }

        private void DataGridViewResults_KeyDown ( object sender, KeyEventArgs e )
        {
            if ( e.Control && e.KeyCode == Keys.A )
            {
                dataGridViewResults.SelectAll();
                e.Handled = true;
            }
        }

        private void ApplyFilter ()
        {
            IEnumerable<ExtensionGroup> groups = _extensionGroups
                .Select( pair => new ExtensionGroup( pair.Key, pair.Value.Count ) );

            string filter = txtExtFilter.Text.Trim();
            if ( !string.IsNullOrEmpty( filter ) )
            {
                groups = groups.Where( g => g.DisplayName.IndexOf( filter, StringComparison.OrdinalIgnoreCase ) >= 0 );
            }

            _filteredGroups = groups
                .OrderBy( g => g.DisplayName, StringComparer.OrdinalIgnoreCase )
                .ToList();

            dataGridViewResults.SuspendLayout();
            dataGridViewResults.Rows.Clear();

            foreach ( ExtensionGroup group in _filteredGroups )
            {
                int rowIndex = dataGridViewResults.Rows.Add( group.DisplayName, group.Count );
                dataGridViewResults.Rows[rowIndex].Tag = group.ExtensionKey;
            }

            dataGridViewResults.ClearSelection();
            dataGridViewResults.ResumeLayout();
            UpdateDeleteButtonState();
        }

        private void UpdateDeleteButtonState ()
        {
            btnDeleteSelected.Enabled = dataGridViewResults.Rows.Count > 0
                && dataGridViewResults.SelectedRows.Count > 0;
        }

        private void btnDeleteSelected_Click ( object sender, EventArgs e )
        {
            if ( dataGridViewResults.SelectedRows.Count == 0 )
            {
                MessageBox.Show( "削除対象の拡張子を選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information );
                return;
            }

            List<string> selectedKeys = dataGridViewResults.SelectedRows
                .Cast<DataGridViewRow>()
                .Select( row => row.Tag as string ?? string.Empty )
                .Distinct( StringComparer.OrdinalIgnoreCase )
                .ToList();

            List<ExtensionDeletionTarget> targets = new List<ExtensionDeletionTarget>();

            foreach ( string key in selectedKeys )
            {
                if ( _extensionGroups.TryGetValue( key, out List<SearchResult> files ) && files.Count > 0 )
                {
                    targets.Add( new ExtensionDeletionTarget( key, files ) );
                }
            }

            if ( targets.Count == 0 )
            {
                MessageBox.Show( "削除可能なファイルがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information );
                return;
            }

            int totalFiles = targets.Sum( t => t.Files.Count );
            if ( totalFiles == 0 )
            {
                MessageBox.Show( "削除可能なファイルがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information );
                return;
            }

            string extensionList = string.Join( ", ", targets.Select( t => t.DisplayName ) );

            DialogResult confirm = MessageBox.Show(
                $"{extensionList} のファイル {totalFiles} 件を削除します。よろしいですか？",
                "削除確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2 );

            if ( confirm != DialogResult.Yes )
            {
                return;
            }

            List<string> errors = new List<string>();

            foreach ( ExtensionDeletionTarget target in targets )
            {
                List<SearchResult> files = target.Files;

                foreach ( SearchResult file in files.ToList() )
                {
                    try
                    {
                        DeleteSingleFile( file.FilePath );

                        if ( files.Remove( file ) && _deletedFileSet.Add( file.FilePath ) )
                        {
                            _deletedFiles.Add( file.FilePath );
                        }
                    }
                    catch ( Exception ex )
                    {
                        errors.Add( $"{file.FilePath}: {ex.Message}" );
                    }
                }

                if ( files.Count == 0 )
                {
                    _extensionGroups.Remove( target.Key );
                }
            }

            if ( errors.Count > 0 )
            {
                MessageBox.Show(
                    string.Join( Environment.NewLine, errors ),
                    "削除エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error );
            }

            ApplyFilter();
        }

        private static void DeleteSingleFile ( string filePath )
        {
            if ( string.IsNullOrWhiteSpace( filePath ) )
            {
                return;
            }

            if ( !File.Exists( filePath ) )
            {
                return;
            }

            FileAttributes attributes = File.GetAttributes( filePath );
            if ( ( attributes & FileAttributes.ReadOnly ) == FileAttributes.ReadOnly )
            {
                File.SetAttributes( filePath, attributes & ~FileAttributes.ReadOnly );
            }

            File.Delete( filePath );
        }

        private class ExtensionGroup
        {
            public ExtensionGroup ( string extensionKey, int count )
            {
                ExtensionKey = extensionKey;
                Count = count;
            }

            public string ExtensionKey { get; }
            public int Count { get; }
            public string DisplayName => string.IsNullOrEmpty( ExtensionKey ) ? NoExtensionDisplayName : ExtensionKey;
        }

        private class ExtensionDeletionTarget
        {
            public ExtensionDeletionTarget ( string key, List<SearchResult> files )
            {
                Key = key;
                Files = files;
            }

            public string Key { get; }
            public List<SearchResult> Files { get; }
            public string DisplayName => string.IsNullOrEmpty( Key ) ? NoExtensionDisplayName : Key;
        }
    }
}
