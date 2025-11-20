using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator
{
    public partial class OtherForm : Form
    {
        private readonly AppSettings _settings;

        public OtherForm(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
            InitializeComponent();
            HookEvents();
        }

        private void HookEvents()
        {
            btnMethodList.Click += BtnMethodList_Click;
            btnCollectFiles.Click += BtnCollectFiles_Click;
        }

        private void BtnMethodList_Click(object sender, EventArgs e)
        {
            using (var dialog = new MethodListForm(_settings))
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ShowDialog(this);
            }
        }

        private void BtnCollectFiles_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new Forms.CollectFilesForm( _settings ) )
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ShowDialog( this );
            }
        }

        private void btnInsertTagJump_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new Forms.InsertTagJumpForm( _settings ) )
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ShowDialog( this );
            }

            SettingsManager.Save( _settings );
        }

        private void btnMethodListWithRowNum_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new MethodListForm( _settings, MethodListExportMode.RowNumber ) )
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ShowDialog( this );
            }
        }

        private void btnInsertTagJumpWithRowNum_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new Forms.InsertTagJumpForm( _settings, TagJumpEmbeddingMode.RowNumber ) )
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ShowDialog( this );
            }

            SettingsManager.Save( _settings );
        }
    }
}
