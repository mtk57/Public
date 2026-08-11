using System.Windows.Forms;

namespace Dir2Txt
{
    public class CloseWithCtrlWForm : Form
    {
        protected override bool ProcessCmdKey ( ref Message msg, Keys keyData )
        {
            if ( keyData == ( Keys.Control | Keys.W ) )
            {
                Close();
                return true;
            }

            return base.ProcessCmdKey( ref msg, keyData );
        }
    }
}
