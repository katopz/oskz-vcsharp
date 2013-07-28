namespace MouseKeyboardActivityMonitor.OSKZ
{
    partial class MainForm
    {
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                m_KeyboardHookManager.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}