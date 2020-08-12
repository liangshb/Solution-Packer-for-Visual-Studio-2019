using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

namespace demo
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class compress
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int compressId = 0x0100;
        public const int sendId = 0x0101;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("73c79c72-6698-440a-88fe-27c609b10c4c");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="compress"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private compress(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCompressID = new CommandID(CommandSet, compressId);
            var menuSendID = new CommandID(CommandSet, sendId);
            var menuItem1 = new MenuCommand(this.compressCallback, menuCompressID);
            var menuItem2 = new MenuCommand(this.sendCallback, menuSendID);
            commandService.AddCommand(menuItem1);
            commandService.AddCommand(menuItem2);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static compress Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in compress's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new compress(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void compressCallback(object sender, EventArgs e)
        {
            // 定位工程目录
            //Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            //DTE dte = (DTE)this.ServiceProvider.GetServiceAsync(typeof(DTE));
            string dirName = ".";
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowNewFolderButton = true;
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
                return;
            else
                dirName = folderBrowserDialog.SelectedPath;

            // 打开对话框选择路径和设置文件名
            string zipFile= "./haha.zip";
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                Filter = "zip压缩文件|*.zip",
                DefaultExt = "zip"
            };
            if (saveFileDialog.ShowDialog() != DialogResult.OK)
                return;
            else
                zipFile = saveFileDialog.FileName;

            // 执行压缩，是否覆盖
            runCompress(zipFile, dirName);

            return;
        }
        private void sendCallback(object sender, EventArgs e)
        {
            VsShellUtilities.ShowMessageBox(this.package, "Todo: send", "To be done",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
        private void runCompress(string zipFile, string dirName)
        {
            try
            {
                File.Delete(zipFile);
                using (ZipStorer zip = ZipStorer.Create(zipFile))
                {
                    zip.AddDirectory(ZipStorer.Compression.Deflate, dirName, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Compress error: " + ex);
            }
            return;
        }
    }
}
