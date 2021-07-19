using System;
using System.EnterpriseServices;
using System.IO;
using System.Reflection;
using System.Windows.Forms.VisualStyles;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Layout;

namespace ExcelExporterImporter
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Button : IExternalApplication
    {
        private const string TabLabel = "BIM One";

        private static readonly ILog Logger =
            LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                var assemblieFolder = Path.GetDirectoryName(Assembly.GetAssembly(GetType()).Location);
                var commandPath = Assembly.GetAssembly(GetType()).Location;

                var fileAppender = new FileAppender {File = assemblieFolder + "\\errors.log", AppendToFile = true};
                var layout = new PatternLayout
                {
                    ConversionPattern = "%date [%thread] %-5level %logger [%property{NDC}] - %message%newline"
                };
                layout.ActivateOptions();

                fileAppender.Layout = layout;
                fileAppender.ActivateOptions();

                BasicConfigurator.Configure(fileAppender);

                var toolsPanel = GetOrCreateRibbonPanel(application);

                PushButton pushButton = toolsPanel.AddItem(new PushButtonData(
                    AddinInfo.ButtonName,
                    AddinInfo.ButtonText,
                    commandPath,
                    "ExcelExporterImporter.Command")) as PushButton;

                var buttonImage = Path.Combine(assemblieFolder, @"Resources\button.png");
                if (!File.Exists(buttonImage))
                    buttonImage = Path.Combine(Directory.GetParent(assemblieFolder).FullName, @"Resources\button.png");

                pushButton.LargeImage = new BitmapImage(new Uri(buttonImage));
                pushButton.ToolTip = AddinInfo.AddinDescription;
            }
            catch (Exception e)
            {
                Logger.Error(e.Message);
                return Result.Failed;
            }


            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private RibbonPanel GetOrCreateRibbonPanel(UIControlledApplication application)
        {
            var ribbonPanel = application.GetRibbonPanels(Tab.AddIns).Find(x => x.Name == TabLabel);
            if (ribbonPanel == null)
                ribbonPanel = application.CreateRibbonPanel(Tab.AddIns, TabLabel);

            return ribbonPanel;
        }
    }
}