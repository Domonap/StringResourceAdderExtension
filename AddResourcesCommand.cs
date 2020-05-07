using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using Task = System.Threading.Tasks.Task;

namespace StringResourceAdderExtension
{
    internal sealed class AddResourcesCommand
    {

        #region Fields
        private static readonly XNamespace X = "http://schemas.microsoft.com/winfx/2006/xaml";
        private static readonly XNamespace XML = "http://www.w3.org/XML/1998/namespace";
        private static readonly XName SPACE = XName.Get("space", XML.NamespaceName);
        private static readonly XName UID = XName.Get("Uid", X.NamespaceName);
        
        private static readonly XName TEXT = XName.Get("Text");
        private static readonly XName HEADER = XName.Get("Header");
        private static readonly XName CONTENT = XName.Get("Content");
        private static readonly XName TOOLTIP = XName.Get("ToolTipService.ToolTip");
        List<XName> HEADERS = new List<XName>() { HEADER, CONTENT, TOOLTIP, TEXT };

        private static readonly XName ROOT = XName.Get("root");
        private static readonly XName NAME = XName.Get("name");
        private static readonly XName VALUE = XName.Get("value");
        private static readonly XName COMMENT = XName.Get("comment");
        private static readonly XName DATA = XName.Get("data");
        List<ProjectItem> projectItems = new List<ProjectItem>();
        int keywordsCount;      
        #endregion Fields 

        #region Command Logic

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("136ad7ed-5e73-474a-bafa-8afd7fd356b0");

        private readonly AsyncPackage package;

        private AddResourcesCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            CommandID menuCommandID = new CommandID(CommandSet, CommandId);
            MenuCommand menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }


        public static AddResourcesCommand Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in SampleCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
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

                keywordsCount = 0;

                var dte = await ServiceProvider.GetServiceAsync(typeof(_DTE)) as _DTE;
                if (dte == null) return;

                var activeDocument = dte.ActiveDocument;
                var source = activeDocument.Object() as TextDocument;
                if (source == null) return;


                var text = source.CreateEditPoint(source.StartPoint).GetText(source.EndPoint);
                var xDocument = XDocument.Parse(text);
                if (xDocument == null) return;


                var keywords = GetKeywords(xDocument);
                if (keywords.Count == 0) return;

                var projects = dte.Solution.Projects.GetEnumerator();
               

                while (projects.MoveNext())
                {
                    var items2 = ((Project)projects.Current).ProjectItems.GetEnumerator();

                    while (items2.MoveNext())
                    {
                        var item2 = (ProjectItem)items2.Current;
                        projectItems.Add(GetFiles(item2));
                    }
                }


                foreach (var file in projectItems)
                    if (file.Name.Contains("Resources.resw"))
                        WriteResources(keywords, file.FileNames[0]);

                ShowMessageBox(
                    $"You added {keywordsCount} strings to Resources",
                    "Resources String Added",
                    OLEMSGICON.OLEMSGICON_INFO);
            }
            catch (Exception)
            {
 
            }
        }

        private  void ShowMessageBox( string message, string title, OLEMSGICON type )
        {
            VsShellUtilities.ShowMessageBox(
                          package,
                         message,
                          title,
                          type,
                          OLEMSGBUTTON.OLEMSGBUTTON_OK,
                          OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

        }

              ProjectItem GetFiles(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //base case
            if (item.ProjectItems == null)
            {
                return item;
            }

            var items = item.ProjectItems.GetEnumerator();
            while (items.MoveNext())
            {
                var currentItem = (ProjectItem)items.Current;
                projectItems.Add(GetFiles(currentItem));
            }

            return item;
        }

        private bool ResourcesExist(XDocument writeFileXml, string resourceName)
        {
            foreach (var node in writeFileXml.Descendants())
            {
                if (node.Name.Equals(DATA))
                {
                    if (node.Attribute(NAME).Value.Equals(resourceName))
                    {
                        
                        return true;
                    }
                }
            }

            return false;
        }
        public  void WriteResources(Dictionary<string, string> keywords, string path)
        {
            try
            {
                var writeFileXml = XDocument.Load(path);

                foreach (KeyValuePair<string, string> item in keywords)
                {
                    if (!ResourcesExist(writeFileXml, item.Key))
                    {
                        var root = new XElement(DATA);
                        root.Add(new XAttribute(NAME, item.Key));
                        root.Add(new XAttribute(SPACE, "preserve"));
                        root.Add(new XElement(VALUE, item.Value));
                        root.Add(new XElement(COMMENT, ""));
                        writeFileXml.Element(ROOT).Add(root);
                       
                    }

                }

                writeFileXml.Save(path);
                 
            }
            catch (Exception e)
            {

                Debug.WriteLine(e.StackTrace);
            }

        }
        public Dictionary<string, string> GetKeywords(XDocument xDocument )
        { 
            Dictionary<string, string> keywords = new Dictionary<string, string>(); 

            foreach (var node in xDocument.Descendants())
            {
                if (node.Attribute(UID) != null)
                {
                    foreach (var attr in node.Attributes())
                    {
                        foreach (var item in HEADERS)
                        {
                            if (attr.Name.Equals(item))
                            {
                                try
                                {
                                    if (node.Attribute(UID).Value.Equals(attr.Value))
                                        ShowMessageBox(
                                                $"x:Uid value in node {node.Name?.LocalName} can't be empty",
                                                $"Missing value",
                                                OLEMSGICON.OLEMSGICON_CRITICAL);
                                   else if (node.Attribute(UID).Value.Equals(attr.Value))
                                        ShowMessageBox(
                                                $"x:Uid value {node.Attribute(UID).Value} can't be the same as the value in {item.LocalName}",
                                                $"Resource {node.Attribute(UID).Value} not added",
                                                OLEMSGICON.OLEMSGICON_WARNING);
                                    else
                                    {
                                        keywords.Add($"{node.Attribute(UID).Value}.{item.LocalName}", attr.Value);
                                        keywordsCount++;
                                    } 
                                }
                                catch (Exception)
                                { 

                                }
                            }
                        } 

                    }

                }

            }


            return keywords;

        }
    }
}
