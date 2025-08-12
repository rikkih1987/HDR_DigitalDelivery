using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class WorksetAssign : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Workset Assign", "This document is not workshared.");
                return Result.Cancelled;
            }

            // pre-run warning (overwrites manual changes)
            var warn = new TaskDialog("HDR_EMEA – Workset Assign")
            {
                MainInstruction = "Assign elements to the most appropriate workset?",
                MainContent = "This will attempt to place elements on the correct workset and may overwrite manual changes.",
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                DefaultButton = TaskDialogResult.No
            };
            if (warn.Show() != TaskDialogResult.Yes)
                return Result.Cancelled;

            // Build workset name → ID lookup
            IList<Workset> userWorksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets();

            Dictionary<string, int> worksetIds = userWorksets
                .ToDictionary(ws => ws.Name, ws => ws.Id.IntegerValue, StringComparer.OrdinalIgnoreCase);

            // Category → workset mapping (clear/known relationships)
            var categoryToWorksetMap = new Dictionary<BuiltInCategory, string>
            {
                { BuiltInCategory.OST_Levels, "QA1_LevelsGrids" },
                { BuiltInCategory.OST_Grids, "QA1_LevelsGrids" },
                { BuiltInCategory.OST_VolumeOfInterest, "QA3_ScopeBoxes" },
                { BuiltInCategory.OST_CLines, "QA3_ReferencePlanes" },
                { BuiltInCategory.OST_RoomSeparationLines, "QA2_SpacesSeparationLines" },
                { BuiltInCategory.OST_MEPSpaces, "QA2_SpacesSeparationLines" },

                // Structural
                { BuiltInCategory.OST_StructuralFoundation, "S1_SubStructure" },
                { BuiltInCategory.OST_StructuralFraming, "S1_SuperStructure" },
                { BuiltInCategory.OST_StructuralColumns, "S1_SuperStructure" },
                { BuiltInCategory.OST_Walls, "S1_SuperStructure" },
                { BuiltInCategory.OST_Floors, "S1_SuperStructure" },
                { BuiltInCategory.OST_Stairs, "S1_SuperStructure" },
                { BuiltInCategory.OST_StairsLandings, "S1_SuperStructure" },
                { BuiltInCategory.OST_Ramps, "S1_SuperStructure" },

                // Electrical containment
                { BuiltInCategory.OST_CableTray, "E1_Containment" },
                { BuiltInCategory.OST_CableTrayFitting, "E1_Containment" },
                { BuiltInCategory.OST_CableTrayRun, "E1_Containment" },
                { BuiltInCategory.OST_Conduit, "E1_Containment" },
                { BuiltInCategory.OST_ConduitFitting, "E1_Containment" },

                // Electrical equipment/devices
                { BuiltInCategory.OST_ElectricalEquipment, "E1_ElectricalEquipment" },
                { BuiltInCategory.OST_DataDevices, "E1_ITTelecoms" },
                { BuiltInCategory.OST_FireAlarmDevices, "E1_FireAlarm" },
                { BuiltInCategory.OST_LightingDevices, "E1_Lighting" },
                { BuiltInCategory.OST_LightingFixtures, "E1_Lighting" },
                { BuiltInCategory.OST_SecurityDevices, "E1_Security" },
                { BuiltInCategory.OST_ElectricalFixtures, "E1_SmallPower" },

                // Mechanical equipment -> Z1_MEPEquipment
                { BuiltInCategory.OST_MechanicalEquipment, "Z1_MEPEquipment" },

                // Fire sprinklers
                { BuiltInCategory.OST_Sprinklers, "F1_SprinklerPipework" },

                // Pipe / Duct systems
                { BuiltInCategory.OST_PipeCurves, "M1_Pipework" },
                { BuiltInCategory.OST_PipeFitting, "M1_Pipework" },
                { BuiltInCategory.OST_PipeAccessory, "M1_Pipework" },

                { BuiltInCategory.OST_DuctCurves, "M1_Ventilation" },
                { BuiltInCategory.OST_DuctFitting, "M1_Ventilation" },
                { BuiltInCategory.OST_DuctAccessory, "M1_Ventilation" },
                { BuiltInCategory.OST_FlexDuctCurves, "M1_Ventilation" },

                // Mass
                { BuiltInCategory.OST_Mass, "Z1_Mass" }
            };

            // Name-pattern → workset mapping (extend over time)
            var namePatternMap = new List<KeyValuePair<string, string[]>>
            {
                // Structure
                new KeyValuePair<string, string[]>("S1_SubStructure",       new[] { "Foundation", "Pile" }),
                new KeyValuePair<string, string[]>("S1_SuperStructure",     new[] { "Framing", "Slab", "BWHRectangular", "Opening", "ColumnHead", "Corbel", "Plinth", "LiftPit", "Joint" }),

                // Containment
                new KeyValuePair<string, string[]>("E1_Containment",        new[] { "ServiceHole", "Conduit", "Tray", "CableCarrier", "CableTray", "Ladder", "Busbar", "BusDuct" }),

                // Electrical
                new KeyValuePair<string, string[]>("E1_ElectricalEquipment",new[] { "AudioVisual", "CommDevice" }),
                new KeyValuePair<string, string[]>("E1_ITTelecoms",         new[] { "CommsAppliance", "DataCabinet", "ONTEnclosure", "FibreGateway" }),
                new KeyValuePair<string, string[]>("E1_FireAlarm",          new[] { "Alarm_Fire", "FireManualCallPoint", "FireIntercom", "FireModule", "FireSwitch", "FireVisualIndicator", "Sensor_Fire", "Aspirating" }),
                new KeyValuePair<string, string[]>("E1_Lighting",           new[] { "LightFixture", "LightingTrack", "ExitSign", "FloodLight", "Bollard", "Controller_Lighting", "LightingControlModule", "LEDDriver", "TrackSpotlight", "NoEntrySign" }),
                new KeyValuePair<string, string[]>("E1_Security",           new[] { "Alarm_Security", "SecurityCamera", "AccessControl", "Intercom", "MagneticLock", "VehicleBarrier", "TrafficLight", "SecurityCorner", "SecurityPanel", "Sensor_Security" }),
                new KeyValuePair<string, string[]>("E1_SmallPower",         new[] { "Outlet_", "PowerWall", "PowerFloor", "PowerCeiling", "PowerBelowDesk", "PowerAboveDesk", "DataGridOutletPoint", "DataFloor", "DataWall", "DataCeiling", "RotaryIsolator", "SwitchIsolator", "EmergencyStop", "CommandoWall", "CarCharging", "ChargingOutlet", "PowerCompartmentFloor", "PowerWithinIntegralTransformer" }),

                // P1 buckets (win before M1)
                new KeyValuePair<string, string[]>("P1_Drainage", new[] {
                    "Drainage", "WasteTerminal", "Gulley", "Gully",
                    "GutterOutlet", "RoofOutlet", "ParapetOutlet", "BalconyOutlet",
                    "Channel", "STrap", "PTrap", "WaterlessTrap", "Tundish",
                    "AirAdmittanceValve", "AAV",
                    "Sewage", "DrainageLiftingStation", "Submersible"
                }),
                new KeyValuePair<string, string[]>("P1_WaterServices", new[] {
                    "ColdWaterStorage", "Calorifier",
                    "ElectricWaterHeater", "ElectricShower", "InstantaneousElectric",
                    "WaterSoftener", "ChlorineDioxide", "DosingPot", "CombinedDosing", "PackagedSideStream",
                    "WaterMeter", "TenantWater", "WaterMeterMultiJet", "WaterMeterUltrasonic"
                }),
                new KeyValuePair<string, string[]>("P1_Condensate", new[] { "Condensate" }),

                // M1 — Ventilation
                new KeyValuePair<string, string[]>("M1_Ventilation", new[] {
                    "AHU", "AirHandling", "CRAC", "CracUnit", "DryAir", "DryAirVBlock", "EvaporativeCooler",
                    "CoolingCoil", "HeatingCoil",
                    "AirTerminal", "Grille", "Louvre", "DiscValve",
                    "DuctSilencer", "DuctFitting", "CapEnd", "Elbow", "Radius", "FlatDuct", "Oval", "RectangularRoundConnection",
                    "Fan_", "Centrifugal", "Axial", "Inline", "Induction", "Impulse", "Destratification", "Hood", "Plug",
                    "Condenser", "HeatPump", "MVHR", "HybridVent", "EnergyMeter"
                }),

                // M1 — Pipework
                new KeyValuePair<string, string[]>("M1_Pipework", new[] {
                    "PipeFitting", "Elbow", "Tee", "Reducer", "Cap", "Union", "Flange",
                    "Valve_", "CheckValve", "BallValve", "ButterflyValve", "GateValve", "Strainer", "Solenoid", "MixingValve",
                    "Thermostatic", "DoubleRegulating", "PressureReducing", "PressureRegulating", "DifferentialPressure", "OrificePlate",
                    "AirRelease", "DrainCock", "SafetyRelief", "SurgeArrestor", "HammerArrestor", "FlushingBypass", "BreechingInlet",
                    "FireFlowSwitch", "WetRiser", "DryRiser",
                    "Pump_", "BoosterSet", "EndSuction", "InlineCirculator", "VerticalInline", "PressurisationUnit",
                    "Tank_", "OilStorage", "ExpansionVessel", "BufferVessel",
                    "HeatExchanger", "HeatInterface",
                    "WaterFilter", "Filter_Water", "HydroMag", "AirSeparator", "DirtSeparator",
                    "FlowInstrument", "PressureGauge", "TemperatureGauge", "FlowSwitch", "Sensor_Temperature", "Sensor_Flow", "TestPoint",
                    "FlowMeter", "GasMeter", "OilMeter"
                }),

                // Sprinklers
                new KeyValuePair<string, string[]>("F1_SprinklerPipework",  new[] { "Sprinkler" }),

                // QA / Z placeholders
                new KeyValuePair<string, string[]>("QA1_LevelsGrids",       new[] { "ExampleLG" }),
                new KeyValuePair<string, string[]>("QA3_ScopeBoxes",        new[] { "ExampleScope" }),
                new KeyValuePair<string, string[]>("QA3_ReferencePlanes",   new[] { "ExampleRef" }),
                new KeyValuePair<string, string[]>("QA2_SpacesSeparationLines", new[] { "ExampleSpace" }),
                new KeyValuePair<string, string[]>("Z1_Mass",               new[] { "ExampleMass" }),

                // Existing (phase override)
                new KeyValuePair<string, string[]>("S2_Existing",           new[] { "ExampleExisting" })
            };

            // ----------------- helpers -----------------

            static bool ContainsIgnoreCase(string haystack, string needle) =>
                !string.IsNullOrEmpty(haystack) &&
                !string.IsNullOrEmpty(needle) &&
                haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;

            static bool AnyTokenMatches(string name, IEnumerable<string> tokens)
            {
                if (string.IsNullOrEmpty(name) || tokens == null) return false;
                foreach (var raw in tokens)
                {
                    if (string.IsNullOrWhiteSpace(raw)) continue;
                    var t = raw.Trim();
                    if (ContainsIgnoreCase(name, t)) return true;
                }
                return false;
            }

            static bool TrySetWorkset(Element e, int targetWsId)
            {
                var p = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                if (p == null || p.IsReadOnly) return false;
                try { p.Set(targetWsId); return true; } catch { return false; }
            }

            // Claim regardless; move if needed
            static bool EnsureWorkset(Element e, int targetWsId, HashSet<ElementId> claimed, HashSet<ElementId> moved)
            {
                claimed.Add(e.Id);
                if (e.WorksetId.IntegerValue == targetWsId) return false;
                if (TrySetWorkset(e, targetWsId)) { moved.Add(e.Id); return true; }
                return false;
            }

            static (string fam, string sym, string typeOrInst, string elem) GetNames(Element e)
            {
                string fam = null, sym = null, typeOrInst = e?.Name, elem = e?.Name;

                var fi = e as FamilyInstance;
                if (fi != null)
                {
                    sym = fi.Symbol?.Name;
                    fam = fi.Symbol?.Family?.Name;
                    typeOrInst = fi.Name; // instance's type display name
                }
                else
                {
                    var et = e?.Document?.GetElement(e.GetTypeId()) as ElementType;
                    if (et != null) sym = et.Name;
                }
                return (fam ?? "", sym ?? "", typeOrInst ?? "", elem ?? "");
            }

            static bool IsPhaseExisting(Element e)
            {
                var p = e.get_Parameter(BuiltInParameter.PHASE_CREATED);
                if (p == null) return false;
                ElementId phId = p.AsElementId();
                if (phId == ElementId.InvalidElementId) return false;
                var phase = e.Document.GetElement(phId) as Phase;
                return phase != null && phase.Name.Equals("Existing", StringComparison.OrdinalIgnoreCase);
            }

            // Only report placed model elements and Revit links
            static bool IsPlacedOrLink(Element e)
            {
                if (e is RevitLinkInstance) return true;
                if (e.Category == null) return false;
                if (e.Category.CategoryType != CategoryType.Model) return false;
                try { return e.get_BoundingBox(null) != null; } catch { return false; }
            }

            static bool IsPipeDomain(Element e)
            {
                int? cat = e.Category != null ? (int?)e.Category.Id.IntegerValue : null;
                return cat == (int)BuiltInCategory.OST_PipeCurves
                    || cat == (int)BuiltInCategory.OST_PipeFitting
                    || cat == (int)BuiltInCategory.OST_PipeAccessory;
            }

            static bool IsDuctDomain(Element e)
            {
                int? cat = e.Category != null ? (int?)e.Category.Id.IntegerValue : null;
                return cat == (int)BuiltInCategory.OST_DuctCurves
                    || cat == (int)BuiltInCategory.OST_DuctFitting
                    || cat == (int)BuiltInCategory.OST_DuctAccessory
                    || cat == (int)BuiltInCategory.OST_FlexDuctCurves;
            }

            // System Classification → workset for pipes/plumbing
            static string MapPipeClassificationToWorkset(Document d, Element e)
            {
                ElementId typeId = ElementId.InvalidElementId;

                var pType = e.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                if (pType != null && pType.StorageType == StorageType.ElementId)
                    typeId = pType.AsElementId();

                var mc = e as MEPCurve;
                if (typeId == ElementId.InvalidElementId && mc != null && mc.MEPSystem != null)
                    typeId = mc.MEPSystem.GetTypeId();

                var pst = d.GetElement(typeId) as PipingSystemType;

                string cls = pst != null ? pst.SystemClassification.ToString() : string.Empty;
                string sysName = pst != null ? pst.Name : string.Empty;

                if (ContainsIgnoreCase(cls, "FireProtection") ||
                    ContainsIgnoreCase(sysName, "Fire") ||
                    ContainsIgnoreCase(sysName, "Sprinkler"))
                    return "F1_SprinklerPipework";

                if (ContainsIgnoreCase(cls, "Sanitary") || ContainsIgnoreCase(cls, "Sewer")
                    || ContainsIgnoreCase(cls, "Drainage") || ContainsIgnoreCase(cls, "Roof"))
                    return "P1_Drainage";

                if (ContainsIgnoreCase(cls, "Domestic") || ContainsIgnoreCase(cls, "ColdWater")
                    || ContainsIgnoreCase(cls, "HotWater") || ContainsIgnoreCase(cls, "Water"))
                    return "P1_WaterServices";

                if (ContainsIgnoreCase(cls, "Condensate") || ContainsIgnoreCase(sysName, "Condensate"))
                    return "P1_Condensate";

                return "M1_Pipework";
            }

            // Find connected Pipes to inherit workset for fittings/accessories
            static IEnumerable<MEPCurve> ConnectedPipes(Element e)
            {
                var results = new HashSet<MEPCurve>();
                ConnectorSet cons = null;

                var mc1 = e as MEPCurve;
                if (mc1 != null && mc1.ConnectorManager != null)
                    cons = mc1.ConnectorManager.Connectors;

                var fi = e as FamilyInstance;
                if (cons == null && fi != null && fi.MEPModel != null && fi.MEPModel.ConnectorManager != null)
                    cons = fi.MEPModel.ConnectorManager.Connectors;

                if (cons == null) yield break;

                foreach (Connector c in cons)
                {
                    foreach (Connector refc in c.AllRefs)
                    {
                        var owner = refc.Owner as Element;
                        var p = owner as Pipe;
                        if (p != null) results.Add(p);
                    }
                }

                foreach (var r in results) yield return r;
            }

            var mechEquipCatId = new ElementId(BuiltInCategory.OST_MechanicalEquipment);

            // Collect all assignable + placed/link elements once
            var allAssignable = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(e =>
                {
                    var p = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                    return p != null && !p.IsReadOnly && IsPlacedOrLink(e);
                })
                .ToList();

            var assignmentCounts = new Dictionary<string, int>();
            var moved = new HashSet<ElementId>();     // changed this run
            var claimed = new HashSet<ElementId>();   // handled by any pass
            var missingTargets = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            Action<string> noteMissing = (wsName) =>
            {
                if (string.IsNullOrEmpty(wsName)) return;
                if (missingTargets.ContainsKey(wsName)) missingTargets[wsName]++;
                else missingTargets[wsName] = 1;
            };

            using (var tx = new Transaction(doc, "Assign elements to worksets"))
            {
                tx.Start();

                // Pass 0: pipes by System Classification
                foreach (var pipe in allAssignable.OfType<Pipe>())
                {
                    string targetWsName = MapPipeClassificationToWorkset(doc, pipe);
                    if (!worksetIds.TryGetValue(targetWsName, out int wsId))
                    {
                        claimed.Add(pipe.Id);
                        noteMissing(targetWsName);
                        continue;
                    }

                    bool changed = EnsureWorkset(pipe, wsId, claimed, moved);
                    if (changed)
                        assignmentCounts[targetWsName] = assignmentCounts.ContainsKey(targetWsName)
                            ? assignmentCounts[targetWsName] + 1 : 1;
                }

                // Pass 0.1: pipe fittings/accessories inherit connected pipe workset (or classification fallback)
                var pipeFA = allAssignable.Where(e => IsPipeDomain(e) && !(e is Pipe)).ToList();
                foreach (var e in pipeFA)
                {
                    int targetWsId = -1;

                    var connectedPipe = ConnectedPipes(e).FirstOrDefault();
                    if (connectedPipe != null)
                    {
                        targetWsId = connectedPipe.WorksetId.IntegerValue;
                        claimed.Add(e.Id);
                    }
                    else
                    {
                        string ws = MapPipeClassificationToWorkset(doc, e);
                        if (worksetIds.TryGetValue(ws, out int tmp))
                            targetWsId = tmp;
                        else
                        {
                            claimed.Add(e.Id);
                            noteMissing(ws);
                        }
                    }

                    if (targetWsId < 0) continue;

                    bool changed = EnsureWorkset(e, targetWsId, claimed, moved);
                    if (changed)
                    {
                        string wsName = userWorksets.FirstOrDefault(w => w.Id.IntegerValue == targetWsId)?.Name ?? "UNKNOWN";
                        assignmentCounts[wsName] = assignmentCounts.ContainsKey(wsName)
                            ? assignmentCounts[wsName] + 1 : 1;
                    }
                }

                // Pass 1: category-based assignment (for everything else)
                foreach (var kvp in categoryToWorksetMap)
                {
                    string targetWorksetName = kvp.Value;
                    if (!worksetIds.TryGetValue(targetWorksetName, out int targetWsId))
                    {
                        var missingElems = new FilteredElementCollector(doc)
                            .OfCategory(kvp.Key)
                            .WhereElementIsNotElementType()
                            .ToElements()
                            .Where(IsPlacedOrLink);
                        foreach (var e in missingElems) claimed.Add(e.Id);
                        noteMissing(targetWorksetName);
                        continue;
                    }

                    var elems = new FilteredElementCollector(doc)
                        .OfCategory(kvp.Key)
                        .WhereElementIsNotElementType()
                        .ToElements()
                        .Where(IsPlacedOrLink)
                        .ToList();

                    int changedCount = 0;
                    foreach (var e in elems)
                        if (!claimed.Contains(e.Id) && EnsureWorkset(e, targetWsId, claimed, moved)) changedCount++;

                    if (changedCount > 0)
                        assignmentCounts[targetWorksetName] = assignmentCounts.ContainsKey(targetWorksetName)
                            ? assignmentCounts[targetWorksetName] + changedCount
                            : changedCount;
                }

                // Pass 2: name-pattern assignment (domain-aware)
                foreach (var e in allAssignable)
                {
                    if (claimed.Contains(e.Id)) continue;

                    // Keep MechanicalEquipment on Z1_MEPEquipment
                    if (e.Category != null && e.Category.Id.IntegerValue == mechEquipCatId.IntegerValue)
                        continue;

                    IEnumerable<KeyValuePair<string, string[]>> buckets = namePatternMap;
                    if (IsPipeDomain(e))
                    {
                        buckets = namePatternMap.Where(k =>
                            k.Key == "P1_Drainage" ||
                            k.Key == "P1_WaterServices" ||
                            k.Key == "P1_Condensate" ||
                            k.Key == "F1_SprinklerPipework" ||
                            k.Key == "M1_Pipework");
                    }
                    else if (IsDuctDomain(e))
                    {
                        buckets = namePatternMap.Where(k => k.Key == "M1_Ventilation");
                    }

                    var names = GetNames(e);
                    string matchedWorkset = null;

                    foreach (var pair in buckets)
                    {
                        var tokens = pair.Value ?? new string[0];
                        if (AnyTokenMatches(names.fam, tokens) ||
                            AnyTokenMatches(names.sym, tokens) ||
                            AnyTokenMatches(names.typeOrInst, tokens) ||
                            AnyTokenMatches(names.elem, tokens))
                        {
                            matchedWorkset = pair.Key;
                            break;
                        }
                    }

                    if (matchedWorkset == null) continue;

                    if (!worksetIds.TryGetValue(matchedWorkset, out int targetWsId2))
                    {
                        claimed.Add(e.Id);
                        noteMissing(matchedWorkset);
                        continue;
                    }

                    if (EnsureWorkset(e, targetWsId2, claimed, moved))
                        assignmentCounts[matchedWorkset] = assignmentCounts.ContainsKey(matchedWorkset)
                            ? assignmentCounts[matchedWorkset] + 1 : 1;
                }

                // Pass 3: phase-based override — "Existing" → S2_Existing
                const string existingName = "S2_Existing";
                if (worksetIds.TryGetValue(existingName, out int existingWsId))
                {
                    var existingElems = allAssignable.Where(IsPhaseExisting).ToList();
                    int movedExisting = 0;
                    foreach (var e in existingElems)
                        if (EnsureWorkset(e, existingWsId, claimed, moved)) movedExisting++;

                    if (movedExisting > 0)
                        assignmentCounts[existingName] = assignmentCounts.ContainsKey(existingName)
                            ? assignmentCounts[existingName] + movedExisting
                            : movedExisting;
                }
                else
                {
                    foreach (var e in allAssignable.Where(IsPhaseExisting)) claimed.Add(e.Id);
                    noteMissing(existingName);
                }

                tx.Commit();
            }

            // Feedback — reassigned counts
            string summary = assignmentCounts.Count == 0
                ? "No elements were reassigned. Ensure worksets exist and categories/tokens are mapped."
                : string.Join("\n", assignmentCounts.Select(kvp => $"{kvp.Value} elements → {kvp.Key}"));

            // Missing targets (we did not create them)
            string missingSummary = "";
            if (missingTargets.Count > 0)
            {
                missingSummary = "\n\nMissing target worksets (no changes applied to those items):\n" +
                    string.Join("\n", missingTargets.OrderBy(k => k.Key)
                        .Select(kvp => $"• {kvp.Key}: {kvp.Value}"));
            }

            // Build skipped set (placed/link + assignable but never claimed)
            var skipped = allAssignable.Where(e => !claimed.Contains(e.Id)).ToList();

            if (skipped.Count > 0)
            {
                // brief label list
                var labelSet = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var e in skipped)
                {
                    var n = GetNames(e);
                    string label = !string.IsNullOrWhiteSpace(n.fam) ? n.fam
                                 : !string.IsNullOrWhiteSpace(n.sym) ? n.sym
                                 : n.elem;
                    if (!string.IsNullOrWhiteSpace(label)) labelSet.Add(label);
                }

                // offer to assign or ignore
                var td = new TaskDialog("HDR_EMEA – Workset Assign — Skipped")
                {
                    MainInstruction = $"Skipped (unmapped/unchanged): {skipped.Count}",
                    MainContent = $"Distinct family/type names: {labelSet.Count}\n\n" +
                                  string.Join("\n", labelSet.Take(25)) +
                                  (labelSet.Count > 25 ? "\n…" : string.Empty)
                };
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Assign all skipped to a workset…");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Ignore skipped items");
                td.CommonButtons = TaskDialogCommonButtons.Close;

                var r = td.Show();
                if (r == TaskDialogResult.CommandLink1)
                {
                    string chosen = PickWorkset(userWorksets.Select(w => w.Name).OrderBy(n => n).ToList());
                    if (!string.IsNullOrWhiteSpace(chosen) && worksetIds.TryGetValue(chosen, out int chosenWsId))
                    {
                        int fixedCount = 0;
                        using (var tx = new Transaction(doc, "Assign skipped to workset"))
                        {
                            tx.Start();
                            foreach (var e in skipped)
                            {
                                var p = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                if (p != null && !p.IsReadOnly)
                                {
                                    try { p.Set(chosenWsId); fixedCount++; } catch { /* ignore */ }
                                }
                            }
                            tx.Commit();
                        }
                        // include in summary
                        if (fixedCount > 0)
                            assignmentCounts[chosen] = assignmentCounts.ContainsKey(chosen)
                                ? assignmentCounts[chosen] + fixedCount
                                : fixedCount;

                        TaskDialog.Show("Workset Assign", $"Assigned {fixedCount} skipped elements to {chosen}.");
                    }
                }
            }

            TaskDialog.Show("Workset Assign", $"Workset assignment complete:\n{summary}{missingSummary}");
            return Result.Succeeded;
        }

        // simple workset picker
        private static string PickWorkset(IList<string> worksetNames)
        {
            using (var form = new WorksetPickerForm(worksetNames))
                return form.ShowDialog() == WinForms.DialogResult.OK ? form.SelectedWorksetName : null;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "WorksetAssign",
                "Workset Assign",
                "HDR_EMEA.Tools_Core.WorksetAssign",
                Properties.Resources.WorksetAssign_32,
                Properties.Resources.WorksetAssign_16,
                "Assigns worksets by category, name tokens, pipe system classification, and connected pipes (placed elements + Revit links)."
            );
        }
    }

    // Minimal WinForms dialog to choose a workset
    internal class WorksetPickerForm : WinForms.Form
    {
        private readonly WinForms.ComboBox _combo = new WinForms.ComboBox();
        private readonly WinForms.Button _ok = new WinForms.Button();
        private readonly WinForms.Button _cancel = new WinForms.Button();

        public string SelectedWorksetName =>
            _combo.SelectedItem == null ? null : _combo.SelectedItem.ToString();

        public WorksetPickerForm(IList<string> worksetNames)
        {
            Text = "Select Workset";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition = WinForms.FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            Width = 380;
            Height = 140;

            _combo.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
            _combo.Left = 12; _combo.Top = 12; _combo.Width = 340;
            foreach (var n in worksetNames) _combo.Items.Add(n);
            if (_combo.Items.Count > 0) _combo.SelectedIndex = 0;

            _ok.Text = "OK"; _ok.DialogResult = WinForms.DialogResult.OK;
            _ok.Left = 192; _ok.Top = 50; _ok.Width = 75;

            _cancel.Text = "Cancel"; _cancel.DialogResult = WinForms.DialogResult.Cancel;
            _cancel.Left = 277; _cancel.Top = 50; _cancel.Width = 75;

            Controls.AddRange(new WinForms.Control[] { _combo, _ok, _cancel });
            AcceptButton = _ok;
            CancelButton = _cancel;
        }
    }
}
