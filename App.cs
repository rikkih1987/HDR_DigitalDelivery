using System;
using System.Diagnostics;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using HDR_EMEA.Tools_Core;
using HDR_EMEA.Common;
using HDR_EMEA.Tools_Structural;

namespace HDR_EMEA
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            const string tabName = "HDR_EMEA";

            // Create tab if missing
            try { app.CreateRibbonTab(tabName); }
            catch { Debug.Print("Tab already exists."); }

            // Create panels
            RibbonPanel panelCoreTools = Utils.CreateRibbonPanel(app, tabName, "Core Tools");
            RibbonPanel panelStruTools = Utils.CreateRibbonPanel(app, tabName, "Structural Tools");

            // --- Core Tools ---
            var btnModelHealth = ModelHealth.GetButtonData();
            var btnSheetManager = SheetManager.GetButtonData();
            var btnModelCoord = ModelCoordinates.GetButtonData();
            var btnAbout = AboutCommand.GetButtonData();
            var btnVisitDDC = VisitDDCCommand.GetButtonData();
            var btnFeedback = FeedbackCommand.GetButtonData();
            var btnScheduleExport = ScheduleExport.GetButtonData();
            var btnFamilyRenamer = FamilyRenamer.GetButtonData();
            var btnRevitDetective = RevitDetective.GetButtonData();
            var btnWorksetAssign = WorksetAssign.GetButtonData();
            var btnWorksetCheck = WorksetCheck.GetButtonData();
            var btnWorksetCreate = WorksetCreate.GetButtonData();
            var btnWorksetRefLink = WorksetRefLink.GetButtonData();
            var btnTagMising = TagMissing.GetButtonData();
            var btnTagDuplicate = TagDuplicates.GetButtonData();
            var btnTextTransform = TextTransform.GetButtonData();
            var btnElementView = ElementView.GetButtonData();
            var btnElevCreator = ElevCreator.GetButtonData();

            // First row:
            panelCoreTools.AddStackedItems(
                btnVisitDDC.Data,
                btnFeedback.Data,
                btnAbout.Data
            );
            // Second row:
            panelCoreTools.AddStackedItems(
                btnModelHealth.Data,
                btnSheetManager.Data,
                btnModelCoord.Data
            );
            // Third row:
            panelCoreTools.AddStackedItems(
                btnScheduleExport.Data,
                btnFamilyRenamer.Data,
                btnRevitDetective.Data
            );
            // Fourth row:
            var worksetSplitData = new SplitButtonData("WorksetSplit", "Assign & Link");
            IList<RibbonItem> worksetItems = panelCoreTools.AddStackedItems(
                btnWorksetCreate.Data,
                worksetSplitData,
                btnWorksetCheck.Data
            );
            if (worksetItems[1] is SplitButton worksetSplit)
            {
                worksetSplit.AddPushButton(btnWorksetAssign.Data);
                worksetSplit.AddPushButton(btnWorksetRefLink.Data);
            }
            // Fifth Row:
            panelCoreTools.AddStackedItems(
                btnTagMising.Data,
                btnTagDuplicate.Data,
                btnTextTransform.Data
            );
            // Sixth Row:
            panelCoreTools.AddStackedItems(
                btnElementView.Data,
                btnElevCreator.Data
            );

            // --- Structural Tools ---
            var btnPileCoord = PileCoordinates.GetButtonData();
            var btnFoundRef = FounRef.GetButtonData();
            var btnPileTagging = PileTagging.GetButtonData();
            var btnProxAdd = PileProximityAdd.GetButtonData();
            var btnProxDel = PileProximityDelete.GetButtonData();
            var btnPileLengths = PileLengths.GetButtonData();
            var btnElementSplitter = ElementSplitter.GetButtonData();

            // First row:
            panelStruTools.AddStackedItems(
                btnPileCoord.Data,
                btnFoundRef.Data,
                btnPileLengths.Data
            );

            // Second row:
            var proxSplitData = new SplitButtonData("PileProxSplit", "Pile Proximity");
            IList<RibbonItem> secondStack = panelStruTools.AddStackedItems(
                proxSplitData,
                btnPileTagging.Data,
                btnElementSplitter.Data
            );
            if (secondStack[0] is SplitButton proxSplit)
            {
                proxSplit.AddPushButton(btnProxAdd.Data);
                proxSplit.AddPushButton(btnProxDel.Data);
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
