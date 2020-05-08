using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using IAsyncServiceProvider = Microsoft.VisualStudio.Shell.IAsyncServiceProvider;
using Task = System.Threading.Tasks.Task;

namespace StringResourceAdderExtension
{
    internal sealed class AddResourcesCommand
    {
        
       #region Fields

        private static readonly XNamespace X = "http://schemas.microsoft.com/winfx/2006/xaml";
        private static readonly XNamespace Xml = "http://www.w3.org/XML/1998/namespace";
        private static readonly XName Space = XName.Get("space", Xml.NamespaceName);
        private static readonly XName Uid = XName.Get("Uid", X.NamespaceName);


        private static readonly XName PlaceholderText = XName.Get("PlaceholderText");
        private static readonly XName Text = XName.Get("Text");
        private static readonly XName Header = XName.Get("Header");
        private static readonly XName Content = XName.Get("Content");
        private static readonly XName Tooltip = XName.Get("ToolTipService.ToolTip");
        private readonly List<XName> _headers = new List<XName> {Header, Content, Tooltip, Text, PlaceholderText};

        private static readonly XName Root = XName.Get("root");
        private static readonly XName Name = XName.Get("name");
        private static readonly XName Value = XName.Get("value");
        private static readonly XName Comment = XName.Get("comment");
        private static readonly XName Data = XName.Get("data");
        private readonly List<ProjectItem> _projectItems = new List<ProjectItem>();
        private int _keywordsCount;

        #endregion Fields

        #region Command Logic

        /// <summary>
        ///     Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        ///     Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("136ad7ed-5e73-474a-bafa-8afd7fd356b0");

        private readonly AsyncPackage _package;

        private AddResourcesCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }


        public static AddResourcesCommand Instance { get; private set; }

        private IAsyncServiceProvider ServiceProvider => _package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in AddResourcesCommand's constructor requires the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new AddResourcesCommand(package, commandService);
        }

        #endregion
        
#pragma warning disable VSTHRD100 // Avoid async void methods
        private async void Execute(object sender, EventArgs e)
#pragma warning restore VSTHRD100 // Avoid async void methods
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                _keywordsCount = 0;

                if (!(await ServiceProvider.GetServiceAsync(typeof(_DTE)) is _DTE dte)) return;

                var activeDocument = dte.ActiveDocument;
                if (!(activeDocument.Object() is TextDocument source)) return;


                var text = source.CreateEditPoint(source.StartPoint).GetText(source.EndPoint);
                var xDocument = XDocument.Parse(text);


                var keywords = GetKeywords(xDocument);
                if (keywords.Count == 0) return;

                var projects = dte.Solution.Projects.GetEnumerator();


                while (projects.MoveNext())
                {
                    var items2 = ((Project) projects.Current)?.ProjectItems.GetEnumerator();

                    while (items2 != null && items2.MoveNext())
                    {
                        var item2 = (ProjectItem) items2.Current;
                        _projectItems.Add(GetFiles(item2));
                    }
                }


                foreach (var file in _projectItems.Where(file =>
                {
                    ThreadHelper.ThrowIfNotOnUIThread();
                    return file.Name.Contains("Resources.resw");
                }))
                    WriteResources(keywords, file.FileNames[0]);


                ShowMessageBox(
                    $"You added {_keywordsCount} strings to Resources",
                    "Resources String Added",
                    OLEMSGICON.OLEMSGICON_INFO);
            }
            catch (Exception)
            {
                ShowMessageBox(
                    "The extension stop working",
                    "Critical error",
                    OLEMSGICON.OLEMSGICON_CRITICAL);
            }
        }


        private ProjectItem GetFiles(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (item.ProjectItems == null) return item;

            var items = item.ProjectItems.GetEnumerator();
            while (items.MoveNext())
            {
                var currentItem = (ProjectItem) items.Current;
                _projectItems.Add(GetFiles(currentItem));
            }

            return item;
        }

        private void ShowMessageBox(string message, string title, OLEMSGICON type)
        {
            VsShellUtilities.ShowMessageBox(
                _package,
                message,
                title,
                type,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private static bool ResourcesExist(XContainer writeFileXml, string resourceName)
        {
            return writeFileXml.Descendants().Where(node => node.Name.Equals(Data))
                .Any(node => node.Attribute(Name).Value.Equals(resourceName));
        }

        private static void WriteResources(Dictionary<string, string> keywords, string path)
        {
            try
            {
                var writeFileXml = XDocument.Load(path);

                foreach (var item in keywords)
                    if (!ResourcesExist(writeFileXml, item.Key))
                    {
                        var root = new XElement(Data);
                        root.Add(new XAttribute(Name, item.Key));
                        root.Add(new XAttribute(Space, "preserve"));
                        root.Add(new XElement(Value, item.Value));
                        root.Add(new XElement(Comment, ""));
                        writeFileXml.Element(Root)?.Add(root);
                    }

                writeFileXml.Save(path);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.StackTrace);
            }
        }

        private Dictionary<string, string> GetKeywords(XContainer xDocument)
        {
            var keywords = new Dictionary<string, string>();

            foreach (var node in xDocument.Descendants())
                if (node.Attribute(Uid) != null)
                    foreach (var attr in node.Attributes())
                    foreach (var item in _headers.Where(item => attr.Name.Equals(item)))
                        try
                        {
                            if (node.Attribute(Uid).Value.Equals(attr.Value))
                            {
                                ShowMessageBox(
                                    $"x:Uid value in node {node.Name?.LocalName} can't be empty",
                                    "Missing value",
                                    OLEMSGICON.OLEMSGICON_CRITICAL);
                            }
                            else if (  node.Attribute(Uid).Value.Equals(attr.Value))
                            {
                                ShowMessageBox(
                                    $"x:Uid value {node.Attribute(Uid)?.Value} can't be the same as the value in {item.LocalName}",
                                    $"Resource {node.Attribute(Uid)?.Value} not added",
                                    OLEMSGICON.OLEMSGICON_WARNING);
                            }
                            else
                            {
                                keywords.Add($"{node.Attribute(Uid)?.Value}.{item.LocalName}", attr.Value);
                                _keywordsCount++;
                            }
                        }
                        catch (Exception)
                        {
                            //ignore
                        }


            return keywords;
        }

 
    }
}