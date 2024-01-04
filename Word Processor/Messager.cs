using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public static class Messager
    {
        public static void ShowErrorMessage(string message, string caption) => MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        public static DialogResult ShowYesNoMessage(string message, string caption) => MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
    }
}