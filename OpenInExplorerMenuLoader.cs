#region Using

using System;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Windows.Input;
using LightPaper.Infrastructure.Contracts;
using Microsoft.Practices.Prism.Logging;
using PuppyFramework;
using PuppyFramework.Interfaces;
using PuppyFramework.MenuService;

#endregion

namespace LightPaper.Plugins.OpenInExplorer
{
    [Export(typeof (IMenuLoader))]
    [Export(typeof (OpenInExplorerMenuLoader))]
    public class OpenInExplorerMenuLoader : IMenuLoader
    {
        #region Fields 

#pragma warning disable 649
        [Import] private Lazy<IMenuFactory> _menuFactory;
        [Import] private Lazy<IMenuRegisterService> _menuRegisterService;
        [Import] private Lazy<IDocumentsManager> _documentsManager;
        private readonly ILogger _logger;
#pragma warning restore 649

        #endregion

        #region Properties 

        public RoutedUICommand OpenInExplorerCommand { get; private set; }

        #endregion

        #region Constructors

        [ImportingConstructor]
        public OpenInExplorerMenuLoader(ILogger logger)
        {
            _logger = logger;
        }

        #endregion

        #region Properties

        public bool IsLoaded { get; private set; }

        #endregion

        #region Methods

        public void Load()
        {
            IsLoaded = true;
            AddMenuItem();
            _logger.Log("Loaded {Type:l}", Category.Info, MagicStrings.PLUGIN_CATEGORY, GetType().Name);
        }

        private void AddMenuItem()
        {
            var openInExplorerGestures = new InputGestureCollection
            {
                new KeyGesture(Key.E, ModifierKeys.Control | ModifierKeys.Alt, "Ctrl+Alt+E")
            };
            OpenInExplorerCommand = new RoutedUICommand("Open in Explorer", "OpenInExplorer", typeof (OpenInExplorerMenuLoader), openInExplorerGestures);

            var fileMenu = _menuFactory.Value.MakeCoreMenuItem(CoreMenuItemType.File);
            var openInExplorerMenu = new MenuItem(OpenInExplorerCommand.Text, 1.31)
            {
                CommandBinding = new CommandBinding(OpenInExplorerCommand, OpenInExplorerCommandExecuted),
            };
            var separator = new SeparatorMenuItem(1.3);
            _menuRegisterService.Value.Register(new MenuItemBase[] {separator, openInExplorerMenu}, fileMenu);
        }

        private void OpenInExplorerCommandExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            var doc = _documentsManager.Value.CurrentDocument;
            if (string.IsNullOrWhiteSpace(doc.SourcePath)) return;
            OpenInExplorer(doc.SourcePath);
        }

        public static void OpenInExplorer(string path)
        {
            if (!File.Exists(path))
            {
                return;
            }

            var argument = string.Format(@"/select, {0}", path);
            Process.Start("explorer.exe", argument);
        }

        #endregion
    }
}