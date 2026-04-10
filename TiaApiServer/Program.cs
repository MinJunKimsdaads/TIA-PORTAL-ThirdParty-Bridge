using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using Siemens.Engineering;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.Hmi.TextGraphicList;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering.SW.Units;
using Siemens.Engineering.MC.Drives;
using Siemens.Engineering.MC.Drives.DFI;

namespace TiaApiServer
{
    class Program
    {
        static string dllDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dll");

        static Program()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string assemblyName = new AssemblyName(args.Name).Name + ".dll";

                // 1) 사용자가 업로드한 dll 폴더 우선 탐색
                string localPath = Path.Combine(dllDir, assemblyName);
                if (File.Exists(localPath))
                    return Assembly.LoadFrom(localPath);

                // 2) TIA Portal 설치 경로 탐색
                string tiaPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    @"Siemens\Automation\Portal V20\PublicAPI\V20");
                string tiaFile = Path.Combine(tiaPath, assemblyName);
                if (File.Exists(tiaFile))
                    return Assembly.LoadFrom(tiaFile);

                // 3) TIA Portal Bin 경로 탐색
                string tiaBinPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    @"Siemens\Automation\Portal V20\Bin\PublicAPI");
                string tiaBinFile = Path.Combine(tiaBinPath, assemblyName);
                if (File.Exists(tiaBinFile))
                    return Assembly.LoadFrom(tiaBinFile);

                return null;
            };
        }

        static JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        static TiaPortal connectedPortal = null;
        static int port = 8099;
        static string tempDir = Path.Combine(Path.GetTempPath(), "TiaApiServer");

        static Dictionary<string, string> langToExt = new Dictionary<string, string>
        {
            { "SCL", ".scl" }, { "DB", ".db" }, { "STL", ".awl" }
        };

        static string MakeValidFileName(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            foreach (char c in invalid) name = name.Replace(c, '_');
            return name;
        }

        static void Main(string[] args)
        {
            if (args.Length > 0) int.TryParse(args[0], out port);

            string prefix = "http://localhost:" + port + "/";
            var listener = new HttpListener();
            listener.Prefixes.Add(prefix);
            listener.Start();

            Console.WriteLine("===========================================");
            Console.WriteLine("  TIA Portal Openness API Server");
            Console.WriteLine("  Running at: " + prefix);
            Console.WriteLine("  Open in browser: " + prefix);
            Console.WriteLine("  Press Ctrl+C to stop");
            Console.WriteLine("===========================================");

            while (true)
            {
                var ctx = listener.GetContext();
                try
                {
                    HandleRequest(ctx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR: " + ex.Message);
                    SendJson(ctx.Response, 500, new { error = ex.Message });
                }
            }
        }

        static void HandleRequest(HttpListenerContext ctx)
        {
            string path = ctx.Request.Url.AbsolutePath.TrimEnd('/');
            string method = ctx.Request.HttpMethod;

            // CORS
            ctx.Response.Headers.Add("Access-Control-Allow-Origin", "*");
            ctx.Response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            ctx.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
            if (method == "OPTIONS") { ctx.Response.StatusCode = 200; ctx.Response.Close(); return; }

            Console.WriteLine(method + " " + path);

            if (path == "" || path == "/") { ServeSwaggerUI(ctx.Response); return; }
            if (path == "/api/openapi.json") { ServeOpenApiSpec(ctx.Response); return; }

            // DLL management
            if (path == "/api/dll/status") { GetDllStatus(ctx.Response); return; }
            if (path == "/api/dll/upload" && method == "POST") { UploadDll(ctx); return; }

            // API routes
            if (path == "/api/processes") { GetProcesses(ctx.Response); return; }
            if (path == "/api/connect") { Connect(ctx); return; }
            if (path == "/api/status") { GetStatus(ctx.Response); return; }
            if (path == "/api/projects") { GetProjects(ctx.Response); return; }
            if (path == "/api/devices") { GetDevices(ctx); return; }
            if (path == "/api/portal/new") { NewPortal(ctx); return; }
            if (path == "/api/project/create") { CreateProject(ctx); return; }
            if (path == "/api/project/add-device") { AddDevice(ctx); return; }
            if (path == "/api/catalog/search") { SearchCatalog(ctx); return; }
            if (path == "/api/catalog/compatible") { SearchCompatible(ctx); return; }
            if (path == "/api/device/items") { GetDeviceItems(ctx); return; }
            if (path == "/api/device/item-attributes") { GetDeviceItemAttributes(ctx); return; }
            if (path == "/api/device/subitems") { GetDeviceSubItems(ctx); return; }
            if (path == "/api/drive/objects") { GetDriveObjects(ctx); return; }
            if (path == "/api/drive/parameters") { GetDriveParameters(ctx); return; }
            if (path == "/api/drive/motor-config") { GetMotorConfiguration(ctx); return; }
            if (path == "/api/drive/dump-all") { DumpAllDriveParameters(ctx); return; }
            if (path.StartsWith("/api/plc/")) { HandlePlcRoute(ctx, path.Substring("/api/plc/".Length)); return; }
            if (path.StartsWith("/api/hmi/")) { HandleHmiRoute(ctx, path.Substring("/api/hmi/".Length)); return; }

            SendJson(ctx.Response, 404, new { error = "Not found", path = path });
        }

        // ── DLL Management ──

        static string[] requiredDlls = new[] { "Siemens.Engineering.dll", "Siemens.Engineering.Hmi.dll" };

        static void GetDllStatus(HttpListenerResponse resp)
        {
            var results = new List<object>();
            foreach (string dllName in requiredDlls)
            {
                string localPath = Path.Combine(dllDir, dllName);
                string tiaPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    @"Siemens\Automation\Portal V20\PublicAPI\V20", dllName);
                string tiaBinPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    @"Siemens\Automation\Portal V20\Bin\PublicAPI", dllName);

                string foundAt = null;
                if (File.Exists(localPath)) foundAt = "dll/" + dllName;
                else if (File.Exists(tiaPath)) foundAt = tiaPath;
                else if (File.Exists(tiaBinPath)) foundAt = tiaBinPath;

                results.Add(new { name = dllName, found = foundAt != null, location = foundAt });
            }

            bool allFound = true;
            foreach (var r in results)
            {
                var foundProp = r.GetType().GetProperty("found");
                if (foundProp != null && !(bool)foundProp.GetValue(r)) { allFound = false; break; }
            }
            SendJson(resp, 200, new {
                ready = allFound,
                message = allFound ? "모든 DLL이 준비되었습니다" : "DLL을 업로드해주세요. TIA Portal이 설치되어 있다면 자동으로 감지됩니다.",
                uploadPath = "dll/",
                dlls = results
            });
        }

        static void UploadDll(HttpListenerContext ctx)
        {
            string fileName = ctx.Request.QueryString["name"];
            if (string.IsNullOrEmpty(fileName) || !fileName.EndsWith(".dll"))
            {
                SendJson(ctx.Response, 400, new { error = "?name= 파라미터에 DLL 파일명을 지정하세요 (예: Siemens.Engineering.dll)" });
                return;
            }

            // 허용된 DLL 이름만 업로드 가능
            if (!fileName.StartsWith("Siemens.Engineering"))
            {
                SendJson(ctx.Response, 400, new { error = "Siemens.Engineering 관련 DLL만 업로드할 수 있습니다" });
                return;
            }

            if (!Directory.Exists(dllDir)) Directory.CreateDirectory(dllDir);

            string savePath = Path.Combine(dllDir, fileName);
            using (var fs = new FileStream(savePath, FileMode.Create))
            {
                ctx.Request.InputStream.CopyTo(fs);
            }

            var fileInfo = new FileInfo(savePath);
            Console.WriteLine("DLL uploaded: " + fileName + " (" + fileInfo.Length + " bytes)");
            SendJson(ctx.Response, 200, new {
                message = "DLL 업로드 완료",
                name = fileName,
                size = fileInfo.Length,
                path = savePath
            });
        }

        // ── TIA Connection ──

        static void GetProcesses(HttpListenerResponse resp)
        {
            var processes = TiaPortal.GetProcesses().Select(p => new {
                id = p.Id,
                projectPath = p.ProjectPath != null ? p.ProjectPath.ToString() : null
            }).ToList();
            SendJson(resp, 200, new { count = processes.Count, processes });
        }

        static void Connect(HttpListenerContext ctx)
        {
            int processId = -1;
            string qid = ctx.Request.QueryString["processId"];
            if (qid != null) int.TryParse(qid, out processId);

            var processes = TiaPortal.GetProcesses().ToList();
            if (processes.Count == 0)
            {
                SendJson(ctx.Response, 404, new { error = "No TIA Portal process found" });
                return;
            }

            TiaPortalProcess target = processId > 0
                ? processes.FirstOrDefault(p => p.Id == processId) ?? processes[0]
                : processes[0];

            connectedPortal = target.Attach();
            SendJson(ctx.Response, 200, new {
                message = "Connected",
                processId = target.Id,
                projectPath = target.ProjectPath != null ? target.ProjectPath.ToString() : null
            });
        }

        static void GetStatus(HttpListenerResponse resp)
        {
            SendJson(resp, 200, new {
                connected = connectedPortal != null,
                projects = connectedPortal != null ? connectedPortal.Projects.Count : 0
            });
        }

        static void EnsureConnected()
        {
            if (connectedPortal != null) return;
            var processes = TiaPortal.GetProcesses().ToList();
            if (processes.Count > 0)
                connectedPortal = processes[0].Attach();
            if (connectedPortal == null)
                throw new Exception("Not connected to TIA Portal. Call /api/connect first.");
        }

        // ── Portal / Project / Device Creation ──

        static void NewPortal(HttpListenerContext ctx)
        {
            string modeStr = ctx.Request.QueryString["mode"] ?? "WithUserInterface";
            TiaPortalMode mode = modeStr == "WithoutUserInterface"
                ? TiaPortalMode.WithoutUserInterface
                : TiaPortalMode.WithUserInterface;

            Console.WriteLine("Starting new TIA Portal (" + mode + ")...");
            var newPortal = new TiaPortal(mode);
            connectedPortal = newPortal;
            SendJson(ctx.Response, 200, new {
                message = "New TIA Portal started and connected",
                mode = mode.ToString()
            });
        }

        static void CreateProject(HttpListenerContext ctx)
        {
            string projectDir = ctx.Request.QueryString["path"];
            string projectName = ctx.Request.QueryString["name"];
            if (string.IsNullOrEmpty(projectDir) || string.IsNullOrEmpty(projectName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?path= and ?name= parameters" });
                return;
            }
            EnsureConnected();

            var dirInfo = new DirectoryInfo(projectDir);
            if (!dirInfo.Exists) dirInfo.Create();

            Console.WriteLine("Creating project: " + projectName + " at " + projectDir);
            var project = connectedPortal.Projects.Create(dirInfo, projectName);
            SendJson(ctx.Response, 200, new {
                message = "Project created",
                name = project.Name,
                path = project.Path != null ? project.Path.ToString() : null
            });
        }

        static void AddDevice(HttpListenerContext ctx)
        {
            string typeId = ctx.Request.QueryString["type"] ?? "System:Device.S71500";
            string deviceName = ctx.Request.QueryString["name"] ?? "NewDevice";
            string orderNumber = ctx.Request.QueryString["order"] ?? "OrderNumber:6ES7 517-3AP00-0AB0/V3.1";
            string projectName = ctx.Request.QueryString["project"];

            EnsureConnected();
            Project project = null;
            foreach (Project p in connectedPortal.Projects)
            {
                if (projectName == null || p.Name == projectName) { project = p; break; }
            }
            if (project == null) { SendJson(ctx.Response, 404, new { error = "No project found" }); return; }

            Console.WriteLine("Adding device: " + deviceName + " type=" + typeId + " order=" + orderNumber);
            var device = project.Devices.CreateWithItem(typeId, deviceName, orderNumber);

            // Find PLC software in the new device
            string plcItemName = null;
            foreach (DeviceItem di in device.DeviceItems)
            {
                var sc = ((IEngineeringServiceProvider)di).GetService<SoftwareContainer>();
                if (sc != null && sc.Software is PlcSoftware)
                {
                    plcItemName = di.Name;
                    break;
                }
            }

            SendJson(ctx.Response, 200, new {
                message = "Device added",
                deviceName = device.Name,
                typeIdentifier = device.TypeIdentifier,
                plcItemName = plcItemName
            });
        }

        // ── Projects & Devices ──

        static void GetProjects(HttpListenerResponse resp)
        {
            EnsureConnected();
            var projects = connectedPortal.Projects.Select(p => new {
                name = p.Name,
                path = p.Path != null ? p.Path.ToString() : null
            }).ToList();
            SendJson(resp, 200, new { count = projects.Count, projects });
        }

        static void GetDevices(HttpListenerContext ctx)
        {
            EnsureConnected();
            string projectName = ctx.Request.QueryString["project"];
            var devices = new List<object>();
            foreach (Project proj in connectedPortal.Projects)
            {
                if (projectName != null && proj.Name != projectName) continue;
                foreach (Device device in proj.Devices)
                {
                    var deviceItems = new List<object>();
                    foreach (DeviceItem di in device.DeviceItems)
                    {
                        string kind = "unknown";
                        var sc = ((IEngineeringServiceProvider)di).GetService<SoftwareContainer>();
                        if (sc != null)
                        {
                            if (sc.Software is PlcSoftware) kind = "PLC";
                            else if (sc.Software is HmiTarget) kind = "HMI";
                        }
                        deviceItems.Add(new {
                            name = di.Name,
                            type = di.TypeIdentifier,
                            softwareKind = kind,
                            hasSoftware = sc != null
                        });
                    }
                    devices.Add(new {
                        name = device.Name,
                        type = device.TypeIdentifier,
                        project = proj.Name,
                        items = deviceItems
                    });
                }
            }
            SendJson(ctx.Response, 200, new { count = devices.Count, devices });
        }

        // ── PLC Routes ──

        static PlcSoftware FindPlc(string deviceItemName)
        {
            EnsureConnected();
            foreach (Project proj in connectedPortal.Projects)
            {
                foreach (Device device in proj.Devices)
                {
                    foreach (DeviceItem di in device.DeviceItems)
                    {
                        if (di.Name == deviceItemName)
                        {
                            var sc = ((IEngineeringServiceProvider)di).GetService<SoftwareContainer>();
                            if (sc != null && sc.Software is PlcSoftware plc) return plc;
                        }
                    }
                }
            }
            return null;
        }

        static void HandlePlcRoute(HttpListenerContext ctx, string route)
        {
            string plcName = ctx.Request.QueryString["plc"];
            if (string.IsNullOrEmpty(plcName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?plc= parameter" });
                return;
            }
            var plc = FindPlc(plcName);
            if (plc == null)
            {
                SendJson(ctx.Response, 404, new { error = "PLC not found: " + plcName });
                return;
            }

            if (route == "blocks") { GetBlocks(ctx.Response, plc.BlockGroup, ""); return; }
            if (route == "types") { GetTypes(ctx.Response, plc); return; }
            if (route == "tags") { GetTagTables(ctx.Response, plc); return; }
            if (route == "units") { GetSoftwareUnits(ctx.Response, plc); return; }
            if (route.StartsWith("units/"))
            {
                string unitPath = route.Substring("units/".Length);
                string unitName = Uri.UnescapeDataString(unitPath.Split('/')[0]);
                string subRoute = unitPath.Contains("/") ? unitPath.Substring(unitPath.IndexOf('/') + 1) : "blocks";

                if (subRoute == "export")
                {
                    ExportUnitZip(ctx.Response, plc, unitName);
                    return;
                }
                GetUnitDetail(ctx.Response, plc, unitName, subRoute);
                return;
            }
            if (route == "block-detail")
            {
                string blockName = ctx.Request.QueryString["name"];
                GetBlockDetail(ctx.Response, plc, blockName);
                return;
            }
            if (route == "export/block")
            {
                string blockName = ctx.Request.QueryString["name"];
                string unitName = ctx.Request.QueryString["unit"];
                ExportSingleBlock(ctx.Response, plc, blockName, unitName);
                return;
            }
            if (route == "export/xml")
            {
                string root = (ctx.Request.QueryString["root"] ?? "all").ToLower();
                ExportAllXmlZip(ctx.Response, plc, root);
                return;
            }
            if (route.StartsWith("units/") && route.EndsWith("/export-xml"))
            {
                string unitPath = route.Substring("units/".Length);
                string uName = Uri.UnescapeDataString(unitPath.Substring(0, unitPath.Length - "/export-xml".Length));
                ExportUnitXmlZip(ctx.Response, plc, uName);
                return;
            }
            if (route == "create-unit")
            {
                string unitName = ctx.Request.QueryString["name"];
                CreateUnit(ctx.Response, plc, unitName);
                return;
            }
            if (route == "import/xml")
            {
                string filePath = ctx.Request.QueryString["file"];
                ImportXml(ctx.Response, plc, filePath);
                return;
            }
            if (route == "import/upload")
            {
                ImportUpload(ctx, plc);
                return;
            }

            SendJson(ctx.Response, 404, new { error = "Unknown PLC route: " + route });
        }

        static void GetBlocks(HttpListenerResponse resp, PlcBlockGroup group, string path)
        {
            var result = CollectBlocks(group, path);
            SendJson(resp, 200, new { count = result.Count, blocks = result });
        }

        static List<object> CollectBlocks(PlcBlockGroup group, string path)
        {
            var result = new List<object>();
            string currentPath = string.IsNullOrEmpty(path) ? group.Name : path + "/" + group.Name;

            foreach (PlcBlock block in group.Blocks)
            {
                result.Add(new {
                    name = block.Name,
                    number = block.Number,
                    type = block.GetType().Name,
                    language = block.ProgrammingLanguage.ToString(),
                    path = currentPath
                });
            }
            foreach (PlcBlockGroup sub in group.Groups)
            {
                result.AddRange(CollectBlocks(sub, currentPath));
            }
            return result;
        }

        static void GetBlockDetail(HttpListenerResponse resp, PlcSoftware plc, string blockName)
        {
            if (string.IsNullOrEmpty(blockName))
            {
                SendJson(resp, 400, new { error = "Missing ?name= parameter" });
                return;
            }
            var block = FindBlock(plc.BlockGroup, blockName);
            if (block == null)
            {
                SendJson(resp, 404, new { error = "Block not found: " + blockName });
                return;
            }
            SendJson(resp, 200, new {
                name = block.Name,
                number = block.Number,
                type = block.GetType().Name,
                language = block.ProgrammingLanguage.ToString(),
                memoryLayout = block.MemoryLayout.ToString(),
                autoNumber = block.AutoNumber
            });
        }

        static PlcBlock FindBlock(PlcBlockGroup group, string name)
        {
            foreach (PlcBlock b in group.Blocks)
                if (b.Name == name) return b;
            foreach (PlcBlockGroup sub in group.Groups)
            {
                var found = FindBlock(sub, name);
                if (found != null) return found;
            }
            return null;
        }

        static void GetTypes(HttpListenerResponse resp, PlcSoftware plc)
        {
            var types = new List<object>();
            foreach (PlcType t in plc.TypeGroup.Types)
            {
                types.Add(new { name = t.Name });
            }
            SendJson(resp, 200, new { count = types.Count, types });
        }

        static void GetTagTables(HttpListenerResponse resp, PlcSoftware plc)
        {
            var tables = new List<object>();
            CollectTagTables(plc.TagTableGroup.TagTables, tables);
            foreach (PlcTagTableUserGroup ug in plc.TagTableGroup.Groups)
                CollectTagTablesFromGroup(ug, tables);
            SendJson(resp, 200, new { count = tables.Count, tagTables = tables });
        }

        static void CollectTagTables(PlcTagTableComposition comp, List<object> result)
        {
            foreach (PlcTagTable table in comp)
            {
                var tags = new List<object>();
                foreach (PlcTag tag in table.Tags)
                {
                    tags.Add(new {
                        name = tag.Name,
                        dataType = tag.DataTypeName,
                        address = tag.LogicalAddress
                    });
                }
                result.Add(new { name = table.Name, tagCount = tags.Count, tags });
            }
        }

        static void CollectTagTablesFromGroup(PlcTagTableUserGroup group, List<object> result)
        {
            CollectTagTables(group.TagTables, result);
            foreach (PlcTagTableUserGroup sub in group.Groups)
                CollectTagTablesFromGroup(sub, result);
        }

        // ── Software Units ──

        static void GetSoftwareUnits(HttpListenerResponse resp, PlcSoftware plc)
        {
            var unitProvider = plc.GetService<PlcUnitProvider>();
            if (unitProvider == null)
            {
                SendJson(resp, 200, new { count = 0, units = new List<object>(), message = "No software units support" });
                return;
            }
            var units = new List<object>();
            foreach (PlcUnit unit in unitProvider.UnitGroup.Units)
            {
                units.Add(new { name = unit.Name, kind = "unit" });
            }
            foreach (PlcSafetyUnit su in unitProvider.UnitGroup.SafetyUnits)
            {
                units.Add(new { name = su.Name, kind = "safety" });
            }
            SendJson(resp, 200, new { count = units.Count, units });
        }

        static void GetUnitDetail(HttpListenerResponse resp, PlcSoftware plc, string unitName, string subRoute)
        {
            var unitProvider = plc.GetService<PlcUnitProvider>();
            if (unitProvider == null) { SendJson(resp, 404, new { error = "No software units" }); return; }

            PlcBlockGroup blockGroup = null;
            PlcTypeSystemGroup typeGroup = null;
            PlcTagTableSystemGroup tagGroup = null;

            foreach (PlcUnit unit in unitProvider.UnitGroup.Units)
            {
                if (unit.Name == unitName)
                {
                    blockGroup = unit.BlockGroup;
                    typeGroup = unit.TypeGroup;
                    tagGroup = unit.TagTableGroup;
                    break;
                }
            }
            if (blockGroup == null)
            {
                foreach (PlcSafetyUnit su in unitProvider.UnitGroup.SafetyUnits)
                {
                    if (su.Name == unitName)
                    {
                        blockGroup = su.BlockGroup;
                        typeGroup = su.TypeGroup;
                        tagGroup = su.TagTableGroup;
                        break;
                    }
                }
            }
            if (blockGroup == null) { SendJson(resp, 404, new { error = "Unit not found: " + unitName }); return; }

            if (subRoute == "blocks")
            {
                var blocks = CollectBlocks(blockGroup, "");
                SendJson(resp, 200, new { unit = unitName, count = blocks.Count, blocks });
            }
            else if (subRoute == "types")
            {
                var types = new List<object>();
                foreach (PlcType t in typeGroup.Types) types.Add(new { name = t.Name });
                SendJson(resp, 200, new { unit = unitName, count = types.Count, types });
            }
            else if (subRoute == "tags")
            {
                var tables = new List<object>();
                CollectTagTables(tagGroup.TagTables, tables);
                SendJson(resp, 200, new { unit = unitName, count = tables.Count, tagTables = tables });
            }
            else
            {
                SendJson(resp, 404, new { error = "Unknown sub-route: " + subRoute });
            }
        }

        // ── Import ──

        static void CreateUnit(HttpListenerResponse resp, PlcSoftware plc, string unitName)
        {
            if (string.IsNullOrEmpty(unitName))
            {
                SendJson(resp, 400, new { error = "Missing ?name= parameter" });
                return;
            }
            var unitProvider = plc.GetService<PlcUnitProvider>();
            if (unitProvider == null)
            {
                SendJson(resp, 500, new { error = "PlcUnitProvider not available" });
                return;
            }
            Console.WriteLine("Creating software unit: " + unitName);
            var unit = unitProvider.UnitGroup.Units.Create(unitName);
            SendJson(resp, 200, new { message = "Unit created", name = unit.Name });
        }

        // ── XML Import (single file or ZIP matching export folder structure) ──

        static void ImportXml(HttpListenerResponse resp, PlcSoftware plc, string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                SendJson(resp, 400, new { error = "Missing ?file= parameter (path to .xml or .zip file)" });
                return;
            }
            if (!File.Exists(filePath))
            {
                SendJson(resp, 404, new { error = "File not found: " + filePath });
                return;
            }

            if (filePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            {
                ImportXmlZip(resp, plc, filePath);
            }
            else
            {
                ImportSingleXmlAuto(resp, plc, filePath);
            }
        }

        static void ImportSingleXmlAuto(HttpListenerResponse resp, PlcSoftware plc, string filePath)
        {
            string content = "";
            try { content = File.ReadAllText(filePath, Encoding.UTF8).Substring(0, Math.Min(2000, (int)new FileInfo(filePath).Length)); } catch { }

            bool isType = content.Contains("<SW.Types.PlcStruct") || content.Contains("<SW.Types.");
            bool isTagTable = content.Contains("<SW.Tags.PlcTagTable") || content.Contains("<SW.Tags.");

            string fileName = Path.GetFileName(filePath);
            string shortPath = CopyToFlat(filePath, @"C:\T\import");

            try
            {
                if (isTagTable)
                {
                    plc.TagTableGroup.TagTables.Import(new FileInfo(shortPath), ImportOptions.Override);
                    SendJson(resp, 200, new { message = "Tag table imported", file = fileName, kind = "tagTable" });
                }
                else if (isType)
                {
                    plc.TypeGroup.Types.Import(new FileInfo(shortPath), ImportOptions.Override);
                    SendJson(resp, 200, new { message = "Type imported", file = fileName, kind = "type" });
                }
                else
                {
                    plc.BlockGroup.Blocks.Import(new FileInfo(shortPath), ImportOptions.Override);
                    SendJson(resp, 200, new { message = "Block imported", file = fileName, kind = "block" });
                }
            }
            catch (Exception ex)
            {
                SendJson(resp, 500, new { error = ex.Message, file = fileName, type = ex.GetType().Name });
            }
            finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
        }

        // ── XML Import via file upload ──

        static void ImportUpload(HttpListenerContext ctx, PlcSoftware plc)
        {
            if (ctx.Request.HttpMethod != "POST")
            {
                SendJson(ctx.Response, 405, new { error = "POST required" });
                return;
            }

            string target = (ctx.Request.QueryString["target"] ?? "blocks").ToLower();   // blocks | tags | types
            string unit = ctx.Request.QueryString["unit"];                                // software unit name (optional)

            // Read uploaded XML body
            byte[] body;
            using (var ms = new MemoryStream())
            {
                ctx.Request.InputStream.CopyTo(ms);
                body = ms.ToArray();
            }
            if (body.Length == 0)
            {
                SendJson(ctx.Response, 400, new { error = "Empty request body. POST the XML file content." });
                return;
            }

            // Save to temp file (Openness API requires FileInfo)
            string tempFile = Path.Combine(tempDir, "upload_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".xml");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            File.WriteAllBytes(tempFile, body);

            // Resolve import target groups
            PlcBlockGroup blockGroup = plc.BlockGroup;
            PlcTypeSystemGroup typeGroup = plc.TypeGroup;
            PlcTagTableSystemGroup tagGroup = plc.TagTableGroup;
            string resolvedUnit = null;

            if (!string.IsNullOrEmpty(unit))
            {
                var unitProvider = plc.GetService<PlcUnitProvider>();
                if (unitProvider == null)
                {
                    File.Delete(tempFile);
                    SendJson(ctx.Response, 500, new { error = "PlcUnitProvider not available" });
                    return;
                }
                bool found = false;
                foreach (PlcUnit pu in unitProvider.UnitGroup.Units)
                {
                    if (pu.Name == unit)
                    {
                        blockGroup = pu.BlockGroup;
                        typeGroup = pu.TypeGroup;
                        tagGroup = pu.TagTableGroup;
                        resolvedUnit = pu.Name;
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    File.Delete(tempFile);
                    SendJson(ctx.Response, 404, new { error = "Software unit not found: " + unit });
                    return;
                }
            }

            // Copy to short path (Openness has path length limits)
            string shortPath = CopyToFlat(tempFile, @"C:\T\import");
            File.Delete(tempFile);

            try
            {
                string fileName = Path.GetFileName(shortPath);
                if (target == "tags")
                {
                    tagGroup.TagTables.Import(new FileInfo(shortPath), ImportOptions.Override);
                    SendJson(ctx.Response, 200, new { message = "Tag table imported", kind = "tags", unit = resolvedUnit });
                }
                else if (target == "types")
                {
                    typeGroup.Types.Import(new FileInfo(shortPath), ImportOptions.Override);
                    SendJson(ctx.Response, 200, new { message = "Type imported", kind = "types", unit = resolvedUnit });
                }
                else
                {
                    blockGroup.Blocks.Import(new FileInfo(shortPath), ImportOptions.Override);
                    SendJson(ctx.Response, 200, new { message = "Block imported", kind = "blocks", unit = resolvedUnit });
                }
            }
            catch (Exception ex)
            {
                SendJson(ctx.Response, 500, new { error = ex.Message, type = ex.GetType().Name, target, unit = resolvedUnit });
            }
            finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
        }

        /// <summary>
        /// Import from ZIP matching export folder structure:
        ///   PLCName/Program blocks/...
        ///   PLCName/PLC data types/...
        ///   PLCName/PLC tags/...
        ///   PLCName/Software units/UnitName/Program blocks/...
        ///   PLCName/Software units/UnitName/PLC data types/...
        ///   PLCName/Software units/UnitName/PLC tags/...
        /// </summary>
        static void ImportXmlZip(HttpListenerResponse resp, PlcSoftware plc, string zipPath)
        {
            var unitProvider = plc.GetService<PlcUnitProvider>();
            string shortBase = @"C:\T\" + DateTime.Now.ToString("HHmmss");
            string extractDir = shortBase + @"\z";
            string flatDir = shortBase + @"\f";
            Directory.CreateDirectory(extractDir);
            Directory.CreateDirectory(flatDir);

            var results = new List<object>();
            int success = 0, failed = 0;

            try
            {
                ZipFile.ExtractToDirectory(zipPath, extractDir);

                // Find root PLC folder (first level in ZIP)
                string[] topDirs = Directory.GetDirectories(extractDir);
                string rootDir = topDirs.Length == 1 ? topDirs[0] : extractDir;

                // ── Phase 1: Root PLC data types (import first for dependency) ──
                string typesDir = Path.Combine(rootDir, "PLC data types");
                if (Directory.Exists(typesDir))
                {
                    Console.WriteLine("=== Importing PLC data types ===");
                    ImportTypesFromXmlDir(typesDir, plc.TypeGroup, null, flatDir, results, ref success, ref failed);
                }

                // ── Phase 2: Root PLC tags ──
                string tagsDir = Path.Combine(rootDir, "PLC tags");
                if (Directory.Exists(tagsDir))
                {
                    Console.WriteLine("=== Importing PLC tags ===");
                    ImportTagTablesFromXmlDir(tagsDir, plc.TagTableGroup, null, flatDir, results, ref success, ref failed);
                }

                // ── Phase 3: Root Program blocks ──
                string blocksDir = Path.Combine(rootDir, "Program blocks");
                if (Directory.Exists(blocksDir))
                {
                    Console.WriteLine("=== Importing Program blocks ===");
                    ImportBlocksFromXmlDir(blocksDir, plc.BlockGroup, null, flatDir, results, ref success, ref failed);
                }

                // ── Phase 4: Software units ──
                // Create units and import ALL types first (cross-unit dependency),
                // then ALL tags, then ALL blocks
                string unitsDir = Path.Combine(rootDir, "Software units");
                var unitMap = new List<KeyValuePair<string, PlcUnit>>();
                if (Directory.Exists(unitsDir) && unitProvider != null)
                {
                    Console.WriteLine("=== Importing Software units ===");
                    foreach (string unitDir in Directory.GetDirectories(unitsDir))
                    {
                        string unitName = Path.GetFileName(unitDir);
                        bool isSafety = unitName.StartsWith("safety_");
                        string actualName = isSafety ? unitName.Substring("safety_".Length) : unitName;

                        PlcUnit targetUnit = null;
                        foreach (PlcUnit u in unitProvider.UnitGroup.Units)
                        {
                            if (u.Name == actualName) { targetUnit = u; break; }
                        }
                        if (targetUnit == null && !isSafety)
                        {
                            try { targetUnit = unitProvider.UnitGroup.Units.Create(actualName); Console.WriteLine("  Created unit: " + actualName); }
                            catch (Exception ex)
                            {
                                results.Add(new { unit = actualName, status = "error", error = "Failed to create unit: " + ex.Message });
                                failed++; continue;
                            }
                        }
                        if (targetUnit == null) { results.Add(new { unit = unitName, status = "skipped", error = "Unit not found" }); continue; }
                        unitMap.Add(new KeyValuePair<string, PlcUnit>(unitDir, targetUnit));
                    }

                    // Phase 4a: All unit tags (before types — types may use tag constants in array bounds)
                    Console.WriteLine("=== Phase 4a: All unit tags ===");
                    foreach (var kv in unitMap)
                    {
                        string uTagsDir = Path.Combine(kv.Key, "PLC tags");
                        if (Directory.Exists(uTagsDir))
                        {
                            Console.WriteLine("--- Tags: " + kv.Value.Name + " ---");
                            ImportTagTablesFromXmlDir(uTagsDir, kv.Value.TagTableGroup, kv.Value.Name, flatDir, results, ref success, ref failed);
                        }
                    }

                    // Phase 4b: All unit types (tag constants now available)
                    Console.WriteLine("=== Phase 4b: All unit types ===");
                    foreach (var kv in unitMap)
                    {
                        string uTypesDir = Path.Combine(kv.Key, "PLC data types");
                        if (Directory.Exists(uTypesDir))
                        {
                            Console.WriteLine("--- Types: " + kv.Value.Name + " ---");
                            ImportTypesFromXmlDir(uTypesDir, kv.Value.TypeGroup, kv.Value.Name, flatDir, results, ref success, ref failed);
                        }
                    }

                    // Phase 4c: All unit blocks
                    Console.WriteLine("=== Phase 4c: All unit blocks ===");
                    foreach (var kv in unitMap)
                    {
                        string uBlocksDir = Path.Combine(kv.Key, "Program blocks");
                        if (Directory.Exists(uBlocksDir))
                        {
                            Console.WriteLine("--- Blocks: " + kv.Value.Name + " ---");
                            ImportBlocksFromXmlDir(uBlocksDir, kv.Value.BlockGroup, kv.Value.Name, flatDir, results, ref success, ref failed);
                        }
                    }
                }

                SendJson(resp, 200, new { message = "XML Import complete", success, failed, details = results });
            }
            finally
            {
                if (Directory.Exists(shortBase)) Directory.Delete(shortBase, true);
            }
        }

        static void ImportTypesFromXmlDir(string dir, PlcTypeSystemGroup typeGroup, string unitName, string flatDir,
            List<object> results, ref int success, ref int failed)
        {
            var pending = Directory.GetFiles(dir, "*.xml", SearchOption.AllDirectories).ToList();
            int maxRounds = 20;

            for (int round = 1; round <= maxRounds && pending.Count > 0; round++)
            {
                var failedThisRound = new List<string>();
                int successBefore = success;

                Console.WriteLine("  Types round " + round + ": " + pending.Count + " files");
                foreach (string xmlFile in pending)
                {
                    string fileName = Path.GetFileName(xmlFile);
                    string shortPath = CopyToFlat(xmlFile, flatDir);
                    try
                    {
                        typeGroup.Types.Import(new FileInfo(shortPath), ImportOptions.Override);
                        Console.WriteLine("  Type OK: " + fileName);
                        results.Add(new { file = fileName, unit = unitName, kind = "type", status = "ok" });
                        success++;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("already exists"))
                        {
                            Console.WriteLine("  Type SKIPPED (exists): " + fileName);
                            results.Add(new { file = fileName, unit = unitName, kind = "type", status = "skipped", error = "Already exists in CPU" });
                            success++;
                        }
                        else
                        {
                            Console.WriteLine("  Type RETRY: " + fileName);
                            failedThisRound.Add(xmlFile);
                        }
                    }
                    finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
                }

                // No progress this round — record remaining as errors and stop
                if (success == successBefore)
                {
                    foreach (string xmlFile in failedThisRound)
                    {
                        string fileName = Path.GetFileName(xmlFile);
                        string shortPath = CopyToFlat(xmlFile, flatDir);
                        try
                        {
                            typeGroup.Types.Import(new FileInfo(shortPath), ImportOptions.Override);
                            results.Add(new { file = fileName, unit = unitName, kind = "type", status = "ok" });
                            success++;
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("already exists"))
                            {
                                Console.WriteLine("  Type SKIPPED (exists): " + fileName);
                                results.Add(new { file = fileName, unit = unitName, kind = "type", status = "skipped", error = "Already exists in CPU" });
                                success++;
                            }
                            else
                            {
                                Console.WriteLine("  Type FAILED: " + fileName + " - " + ex.Message);
                                results.Add(new { file = fileName, unit = unitName, kind = "type", status = "error", error = ex.Message });
                                failed++;
                            }
                        }
                        finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
                    }
                    break;
                }

                pending = failedThisRound;
            }
        }

        static void ImportTagTablesFromXmlDir(string dir, PlcTagTableSystemGroup tagGroup, string unitName, string flatDir,
            List<object> results, ref int success, ref int failed)
        {
            foreach (string xmlFile in Directory.GetFiles(dir, "*.xml", SearchOption.AllDirectories))
            {
                string fileName = Path.GetFileName(xmlFile);
                string shortPath = CopyToFlat(xmlFile, flatDir);
                Console.WriteLine("  TagTable: " + fileName + (unitName != null ? " -> " + unitName : " -> root"));
                try
                {
                    tagGroup.TagTables.Import(new FileInfo(shortPath), ImportOptions.Override);
                    results.Add(new { file = fileName, unit = unitName, kind = "tagTable", status = "ok" });
                    success++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  FAILED: " + ex.Message);
                    results.Add(new { file = fileName, unit = unitName, kind = "tagTable", status = "error", error = ex.Message });
                    failed++;
                }
                finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
            }
        }

        static void ImportBlocksFromXmlDir(string dir, PlcBlockGroup blockGroup, string unitName, string flatDir,
            List<object> results, ref int success, ref int failed)
        {
            // Import order: non-instance blocks first, then DB/instance DBs
            var pending = Directory.GetFiles(dir, "*.xml", SearchOption.AllDirectories)
                .OrderBy(f =>
                {
                    string fn = f.ToLower();
                    if (fn.Contains("instance db")) return 2;
                    if (fn.EndsWith("_db.xml")) return 1;
                    return 0;
                })
                .ToList();

            int maxRounds = 20;

            for (int round = 1; round <= maxRounds && pending.Count > 0; round++)
            {
                var failedThisRound = new List<string>();
                int successBefore = success;

                Console.WriteLine("  Blocks round " + round + ": " + pending.Count + " files");
                foreach (string xmlFile in pending)
                {
                    string fileName = Path.GetFileName(xmlFile);
                    string shortPath = CopyToFlat(xmlFile, flatDir);

                    string relDir = Path.GetDirectoryName(xmlFile);
                    string relativePath = relDir.Length > dir.Length ? relDir.Substring(dir.Length).TrimStart('\\', '/') : "";

                    PlcBlockGroup targetGroup = NavigateOrCreateBlockGroup(blockGroup, relativePath);

                    try
                    {
                        targetGroup.Blocks.Import(new FileInfo(shortPath), ImportOptions.Override);
                        Console.WriteLine("  Block OK: " + fileName);
                        results.Add(new { file = fileName, unit = unitName, path = relativePath, kind = "block", status = "ok" });
                        success++;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("at least one compile unit") || ex.Message.Contains("already exists"))
                        {
                            string skipReason = ex.Message.Contains("already exists") ? "Already exists in CPU" : "Library block (no compile unit)";
                            Console.WriteLine("  Block SKIPPED: " + fileName + " - " + skipReason);
                            results.Add(new { file = fileName, unit = unitName, path = relativePath, kind = "block", status = "skipped", error = skipReason });
                            success++;
                        }
                        else
                        {
                            Console.WriteLine("  Block RETRY: " + fileName);
                            failedThisRound.Add(xmlFile);
                        }
                    }
                    finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
                }

                // No progress — record remaining as errors
                if (success == successBefore)
                {
                    foreach (string xmlFile in failedThisRound)
                    {
                        string fileName = Path.GetFileName(xmlFile);
                        string shortPath = CopyToFlat(xmlFile, flatDir);
                        string relDir = Path.GetDirectoryName(xmlFile);
                        string relativePath = relDir.Length > dir.Length ? relDir.Substring(dir.Length).TrimStart('\\', '/') : "";
                        PlcBlockGroup targetGroup = NavigateOrCreateBlockGroup(blockGroup, relativePath);

                        try
                        {
                            targetGroup.Blocks.Import(new FileInfo(shortPath), ImportOptions.Override);
                            results.Add(new { file = fileName, unit = unitName, path = relativePath, kind = "block", status = "ok" });
                            success++;
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("at least one compile unit") || ex.Message.Contains("already exists"))
                            {
                                string skipReason = ex.Message.Contains("already exists") ? "Already exists in CPU" : "Library block (no compile unit)";
                                Console.WriteLine("  Block SKIPPED: " + fileName + " - " + skipReason);
                                results.Add(new { file = fileName, unit = unitName, path = relativePath, kind = "block", status = "skipped", error = skipReason });
                                success++;
                            }
                            else
                            {
                                Console.WriteLine("  Block FAILED: " + fileName + " - " + ex.Message);
                                results.Add(new { file = fileName, unit = unitName, path = relativePath, kind = "block", status = "error", error = ex.Message });
                                failed++;
                            }
                        }
                        finally { if (File.Exists(shortPath)) File.Delete(shortPath); }
                    }
                    break;
                }

                pending = failedThisRound;
            }
        }

        static PlcBlockGroup NavigateOrCreateBlockGroup(PlcBlockGroup root, string relativePath)
        {
            if (string.IsNullOrEmpty(relativePath)) return root;

            PlcBlockGroup current = root;
            string[] parts = relativePath.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                PlcBlockGroup found = null;
                foreach (PlcBlockGroup g in current.Groups)
                {
                    if (g.Name == part) { found = g; break; }
                }
                if (found == null)
                {
                    try { found = current.Groups.Create(part); Console.WriteLine("  Created group: " + part); }
                    catch { break; }
                }
                current = found;
            }
            return current;
        }

        // ── XML Export (all block types including LAD/FBD) ──

        static void ExportBlocksXmlToZip(PlcBlockGroup group, string currentPath, string zipPrefix, ZipArchive zip, string exportDir, List<Dictionary<string, object>> results = null)
        {
            // currentPath is empty for the root group (skip root group name since prefix already has it)
            foreach (PlcBlock block in group.Blocks)
            {
                string fileName = MakeValidFileName(block.Name) + ".xml";
                string tempFile = Path.Combine(exportDir, fileName);
                try
                {
                    if (File.Exists(tempFile)) File.Delete(tempFile);
                    block.Export(new FileInfo(tempFile), ExportOptions.WithDefaults);
                    if (File.Exists(tempFile))
                    {
                        string entryPath = string.IsNullOrEmpty(currentPath)
                            ? zipPrefix + fileName
                            : zipPrefix + currentPath + "/" + fileName;
                        zip.CreateEntryFromFile(tempFile, entryPath);
                        File.Delete(tempFile);
                        Console.WriteLine("  XML: " + entryPath);
                        if (results != null) results.Add(new Dictionary<string, object> { {"file", fileName}, {"path", currentPath}, {"kind", "block"}, {"status", "ok"} });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  XML export error [" + block.Name + "]: " + ex.Message);
                    if (results != null) results.Add(new Dictionary<string, object> { {"file", fileName}, {"path", currentPath}, {"kind", "block"}, {"status", "error"}, {"error", ex.Message} });
                }
            }
            foreach (PlcBlockGroup sub in group.Groups)
            {
                string subPath = string.IsNullOrEmpty(currentPath) ? sub.Name : currentPath + "/" + sub.Name;
                ExportBlocksXmlToZip(sub, subPath, zipPrefix, zip, exportDir, results);
            }
        }

        static void ExportTypesXmlToZip(PlcTypeSystemGroup typeGroup, string zipPrefix, ZipArchive zip, string exportDir, List<Dictionary<string, object>> results = null)
        {
            ExportTypeGroupXmlToZip(typeGroup.Types, typeGroup.Groups, "", zipPrefix, zip, exportDir, results);
        }

        static void ExportTypeGroupXmlToZip(PlcTypeComposition types, PlcTypeUserGroupComposition groups, string groupPath, string zipPrefix, ZipArchive zip, string exportDir, List<Dictionary<string, object>> results = null)
        {
            foreach (PlcType t in types)
            {
                string fileName = MakeValidFileName(t.Name) + ".xml";
                string tempFile = Path.Combine(exportDir, fileName);
                try
                {
                    if (File.Exists(tempFile)) File.Delete(tempFile);
                    t.Export(new FileInfo(tempFile), ExportOptions.WithDefaults);
                    if (File.Exists(tempFile))
                    {
                        string entryPath = zipPrefix + groupPath + fileName;
                        zip.CreateEntryFromFile(tempFile, entryPath);
                        File.Delete(tempFile);
                        Console.WriteLine("  XML: " + entryPath);
                        if (results != null) results.Add(new Dictionary<string, object> { {"file", fileName}, {"path", groupPath.TrimEnd('/')}, {"kind", "type"}, {"status", "ok"} });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  XML type export error [" + t.Name + "]: " + ex.Message);
                    if (results != null) results.Add(new Dictionary<string, object> { {"file", fileName}, {"path", groupPath.TrimEnd('/')}, {"kind", "type"}, {"status", "error"}, {"error", ex.Message} });
                }
            }
            foreach (PlcTypeUserGroup sub in groups)
            {
                string subPath = groupPath + MakeValidFileName(sub.Name) + "/";
                ExportTypeGroupXmlToZip(sub.Types, sub.Groups, subPath, zipPrefix, zip, exportDir, results);
            }
        }

        static void ExportUnitXmlZip(HttpListenerResponse resp, PlcSoftware plc, string unitName)
        {
            var unitProvider = plc.GetService<PlcUnitProvider>();
            if (unitProvider == null) { SendJson(resp, 404, new { error = "No software units" }); return; }

            PlcBlockGroup blockGroup = null;
            PlcTypeSystemGroup typeGroup = null;
            PlcTagTableSystemGroup tagGroup = null;

            foreach (PlcUnit unit in unitProvider.UnitGroup.Units)
            {
                if (unit.Name == unitName) { blockGroup = unit.BlockGroup; typeGroup = unit.TypeGroup; tagGroup = unit.TagTableGroup; break; }
            }
            if (blockGroup == null)
            {
                foreach (PlcSafetyUnit su in unitProvider.UnitGroup.SafetyUnits)
                {
                    if (su.Name == unitName) { blockGroup = su.BlockGroup; typeGroup = su.TypeGroup; tagGroup = su.TagTableGroup; break; }
                }
            }
            if (blockGroup == null) { SendJson(resp, 404, new { error = "Unit not found: " + unitName }); return; }

            string exportDir = Path.Combine(tempDir, Guid.NewGuid().ToString().Substring(0, 8));
            Directory.CreateDirectory(exportDir);
            string zipFile = Path.Combine(tempDir, MakeValidFileName(unitName) + "_xml.zip");
            if (File.Exists(zipFile)) File.Delete(zipFile);

            Console.WriteLine("XML exporting unit: " + unitName);
            try
            {
                using (var zip = ZipFile.Open(zipFile, ZipArchiveMode.Create))
                {
                    string prefix = MakeValidFileName(unitName) + "/";
                    ExportBlocksXmlToZip(blockGroup, "", prefix, zip, exportDir);
                    ExportTypesXmlToZip(typeGroup, prefix, zip, exportDir);
                    ExportTagTablesToZip(tagGroup.TagTables, prefix, zip, exportDir, tagGroup.Groups);
                }
                SendFile(resp, zipFile, MakeValidFileName(unitName) + "_xml.zip");
            }
            finally
            {
                if (File.Exists(zipFile)) File.Delete(zipFile);
                if (Directory.Exists(exportDir)) Directory.Delete(exportDir, true);
            }
        }

        static void ExportAllXmlZip(HttpListenerResponse resp, PlcSoftware plc, string root = "all")
        {
            if (root != "all" && root != "program_blocks" && root != "software_units")
            {
                SendJson(resp, 400, new { error = "Invalid root parameter. Use: all, program_blocks, software_units" });
                return;
            }

            string plcName = MakeValidFileName(plc.Name);
            string exportDir = Path.Combine(tempDir, Guid.NewGuid().ToString().Substring(0, 8));
            Directory.CreateDirectory(exportDir);
            string zipFile = Path.Combine(tempDir, plcName + "_xml_" + root + ".zip");
            if (File.Exists(zipFile)) File.Delete(zipFile);

            Console.WriteLine("XML exporting from PLC: " + plc.Name + " (root=" + root + ")");
            var results = new List<Dictionary<string, object>>();
            try
            {
                using (var zip = ZipFile.Open(zipFile, ZipArchiveMode.Create))
                {
                    string plcPrefix = plcName + "/";

                    // Program blocks
                    if (root == "all" || root == "program_blocks")
                    {
                        ExportBlocksXmlToZip(plc.BlockGroup, "", plcPrefix + "Program blocks/", zip, exportDir, results);
                    }

                    // PLC data types
                    if (root == "all" || root == "program_blocks")
                    {
                        ExportTypesXmlToZip(plc.TypeGroup, plcPrefix + "PLC data types/", zip, exportDir, results);
                    }

                    // PLC tags
                    if (root == "all" || root == "program_blocks")
                    {
                        ExportTagTablesToZip(plc.TagTableGroup.TagTables, plcPrefix + "PLC tags/", zip, exportDir, plc.TagTableGroup.Groups, results);
                    }

                    // Software units
                    if (root == "all" || root == "software_units")
                    {
                        var unitProvider = plc.GetService<PlcUnitProvider>();
                        if (unitProvider != null)
                        {
                            string suBase = plcPrefix + "Software units/";
                            foreach (PlcUnit unit in unitProvider.UnitGroup.Units)
                            {
                                string unitPrefix = suBase + MakeValidFileName(unit.Name) + "/";
                                Console.WriteLine("  XML exporting unit: " + unit.Name);
                                ExportBlocksXmlToZip(unit.BlockGroup, "", unitPrefix + "Program blocks/", zip, exportDir, results);
                                ExportTypesXmlToZip(unit.TypeGroup, unitPrefix + "PLC data types/", zip, exportDir, results);
                                ExportTagTablesToZip(unit.TagTableGroup.TagTables, unitPrefix + "PLC tags/", zip, exportDir, unit.TagTableGroup.Groups, results);
                            }
                            foreach (PlcSafetyUnit su in unitProvider.UnitGroup.SafetyUnits)
                            {
                                string unitPrefix = suBase + "safety_" + MakeValidFileName(su.Name) + "/";
                                Console.WriteLine("  XML exporting safety unit: " + su.Name);
                                ExportBlocksXmlToZip(su.BlockGroup, "", unitPrefix + "Program blocks/", zip, exportDir, results);
                                ExportTypesXmlToZip(su.TypeGroup, unitPrefix + "PLC data types/", zip, exportDir, results);
                                ExportTagTablesToZip(su.TagTableGroup.TagTables, unitPrefix + "PLC tags/", zip, exportDir, su.TagTableGroup.Groups, results);
                            }
                        }
                    }

                    // Add export report inside ZIP
                    int okCount = results.Count(r => (string)r["status"] == "ok");
                    int errCount = results.Count(r => (string)r["status"] == "error");
                    var errorDetails = results.Where(r => (string)r["status"] == "error").ToList();
                    var report = new Dictionary<string, object> {
                        {"message", "XML Export complete"},
                        {"success", okCount},
                        {"failed", errCount},
                        {"details", errorDetails}
                    };
                    string reportJson = json.Serialize(report);
                    var reportEntry = zip.CreateEntry(plcName + "/_export_report.json");
                    using (var sw = new StreamWriter(reportEntry.Open())) { sw.Write(reportJson); }
                }

                int totalOk = results.Count(r => (string)r["status"] == "ok");
                int totalErr = results.Count(r => (string)r["status"] == "error");
                SendFileWithExportReport(resp, zipFile, plcName + "_xml_" + root + ".zip", totalOk, totalErr);
            }
            finally
            {
                if (File.Exists(zipFile)) File.Delete(zipFile);
                if (Directory.Exists(exportDir)) Directory.Delete(exportDir, true);
            }
        }

        // ── XML Import ──

        static string CopyToFlat(string sourceFile, string flatDir)
        {
            if (!Directory.Exists(flatDir)) Directory.CreateDirectory(flatDir);
            string fileName = Path.GetFileName(sourceFile);
            string shortPath = Path.Combine(flatDir, fileName);
            int counter = 1;
            while (File.Exists(shortPath))
            {
                shortPath = Path.Combine(flatDir, Path.GetFileNameWithoutExtension(fileName) + "_" + counter + Path.GetExtension(fileName));
                counter++;
            }
            File.Copy(sourceFile, shortPath);
            return shortPath;
        }

        // ── Export / Download ──

        static string ExportBlockToFile(PlcBlock block, PlcExternalSourceSystemGroup extGroup, string outputDir)
        {
            string lang = block.ProgrammingLanguage.ToString();
            if (!langToExt.ContainsKey(lang)) return null;

            string ext = langToExt[lang];
            string fileName = MakeValidFileName(block.Name) + ext;
            string tempFile = Path.Combine(outputDir, fileName);
            var fileInfo = new FileInfo(tempFile);

            try
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
                extGroup.GenerateSource(new List<PlcBlock> { block }, fileInfo, GenerateOptions.None);
                return tempFile;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Export error [" + block.Name + "]: " + ex.Message);
                return null;
            }
        }

        static void ExportBlocksToZip(PlcBlockGroup group, PlcExternalSourceSystemGroup extGroup,
            string currentPath, string zipPrefix, ZipArchive zip, string exportDir)
        {
            string groupPath = string.IsNullOrEmpty(currentPath) ? group.Name : currentPath + "/" + group.Name;

            foreach (PlcBlock block in group.Blocks)
            {
                string lang = block.ProgrammingLanguage.ToString();
                if (!langToExt.ContainsKey(lang)) continue;

                string exported = ExportBlockToFile(block, extGroup, exportDir);
                if (exported != null && File.Exists(exported))
                {
                    string langFolder = lang.ToLower();
                    string entryPath = zipPrefix + langFolder + "/" + groupPath + "/" + MakeValidFileName(block.Name) + langToExt[lang];
                    zip.CreateEntryFromFile(exported, entryPath);
                    File.Delete(exported);
                    Console.WriteLine("  Zipped: " + entryPath);
                }
            }
            foreach (PlcBlockGroup sub in group.Groups)
            {
                ExportBlocksToZip(sub, extGroup, groupPath, zipPrefix, zip, exportDir);
            }
        }

        static void ExportTypesToZip(PlcTypeSystemGroup typeGroup, PlcExternalSourceSystemGroup extGroup,
            string zipPrefix, ZipArchive zip, string exportDir)
        {
            foreach (PlcType t in typeGroup.Types)
            {
                string fileName = t.Name + ".udt";
                string tempFile = Path.Combine(exportDir, fileName);
                var fileInfo = new FileInfo(tempFile);
                try
                {
                    if (File.Exists(tempFile)) File.Delete(tempFile);
                    extGroup.GenerateSource(new List<PlcType> { t }, fileInfo, GenerateOptions.None);
                    if (File.Exists(tempFile))
                    {
                        string entryPath = zipPrefix + "udt/" + fileName;
                        zip.CreateEntryFromFile(tempFile, entryPath);
                        File.Delete(tempFile);
                        Console.WriteLine("  Zipped: " + entryPath);
                    }
                }
                catch (Exception ex) { Console.WriteLine("Type export error [" + t.Name + "]: " + ex.Message); }
            }
        }

        static void ExportTagTablesToZip(PlcTagTableComposition tables, string zipPrefix, ZipArchive zip, string exportDir, PlcTagTableUserGroupComposition groups = null, List<Dictionary<string, object>> results = null)
        {
            ExportTagTableGroupToZip(tables, groups, "", zipPrefix, zip, exportDir, results);
        }

        static void ExportTagTableGroupToZip(PlcTagTableComposition tables, PlcTagTableUserGroupComposition groups, string groupPath, string zipPrefix, ZipArchive zip, string exportDir, List<Dictionary<string, object>> results = null)
        {
            foreach (PlcTagTable table in tables)
            {
                string fileName = MakeValidFileName(table.Name) + ".xml";
                string tempFile = Path.Combine(exportDir, fileName);
                try
                {
                    if (File.Exists(tempFile)) File.Delete(tempFile);
                    table.Export(new FileInfo(tempFile), ExportOptions.WithDefaults);
                    if (File.Exists(tempFile))
                    {
                        string entryPath = zipPrefix + groupPath + fileName;
                        zip.CreateEntryFromFile(tempFile, entryPath);
                        File.Delete(tempFile);
                        Console.WriteLine("  Zipped: " + entryPath);
                        if (results != null) results.Add(new Dictionary<string, object> { {"file", fileName}, {"path", groupPath.TrimEnd('/')}, {"kind", "tagTable"}, {"status", "ok"} });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Tag export error [" + table.Name + "]: " + ex.Message);
                    if (results != null) results.Add(new Dictionary<string, object> { {"file", fileName}, {"path", groupPath.TrimEnd('/')}, {"kind", "tagTable"}, {"status", "error"}, {"error", ex.Message} });
                }
            }
            if (groups != null)
            {
                foreach (PlcTagTableUserGroup sub in groups)
                {
                    string subPath = groupPath + MakeValidFileName(sub.Name) + "/";
                    ExportTagTableGroupToZip(sub.TagTables, sub.Groups, subPath, zipPrefix, zip, exportDir, results);
                }
            }
        }

        static void ExportSingleBlock(HttpListenerResponse resp, PlcSoftware plc, string blockName, string unitName)
        {
            if (string.IsNullOrEmpty(blockName))
            {
                SendJson(resp, 400, new { error = "Missing ?name= parameter" });
                return;
            }

            PlcBlock block = null;
            PlcExternalSourceSystemGroup extGroup = null;

            if (string.IsNullOrEmpty(unitName))
            {
                block = FindBlock(plc.BlockGroup, blockName);
                extGroup = plc.ExternalSourceGroup;
            }
            else
            {
                var unitProvider = plc.GetService<PlcUnitProvider>();
                if (unitProvider != null)
                {
                    foreach (PlcUnit unit in unitProvider.UnitGroup.Units)
                    {
                        if (unit.Name == unitName)
                        {
                            block = FindBlock(unit.BlockGroup, blockName);
                            extGroup = unit.ExternalSourceGroup;
                            break;
                        }
                    }
                    if (block == null)
                    {
                        foreach (PlcSafetyUnit su in unitProvider.UnitGroup.SafetyUnits)
                        {
                            if (su.Name == unitName)
                            {
                                block = FindBlock(su.BlockGroup, blockName);
                                extGroup = su.ExternalSourceGroup;
                                break;
                            }
                        }
                    }
                }
            }

            if (block == null) { SendJson(resp, 404, new { error = "Block not found: " + blockName }); return; }
            if (extGroup == null) { SendJson(resp, 500, new { error = "ExternalSourceGroup not available" }); return; }

            string lang = block.ProgrammingLanguage.ToString();
            if (!langToExt.ContainsKey(lang))
            {
                SendJson(resp, 400, new { error = "Language not exportable: " + lang + " (only SCL, DB, STL supported)" });
                return;
            }

            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            string exported = ExportBlockToFile(block, extGroup, tempDir);
            if (exported == null || !File.Exists(exported))
            {
                SendJson(resp, 500, new { error = "Export failed for block: " + blockName });
                return;
            }

            SendFile(resp, exported, MakeValidFileName(block.Name) + langToExt[lang]);
            File.Delete(exported);
        }

        static void ExportUnitZip(HttpListenerResponse resp, PlcSoftware plc, string unitName)
        {
            var unitProvider = plc.GetService<PlcUnitProvider>();
            if (unitProvider == null) { SendJson(resp, 404, new { error = "No software units" }); return; }

            PlcBlockGroup blockGroup = null;
            PlcExternalSourceSystemGroup extGroup = null;
            PlcTypeSystemGroup typeGroup = null;
            PlcTagTableSystemGroup tagGroup = null;

            foreach (PlcUnit unit in unitProvider.UnitGroup.Units)
            {
                if (unit.Name == unitName)
                {
                    blockGroup = unit.BlockGroup;
                    extGroup = unit.ExternalSourceGroup;
                    typeGroup = unit.TypeGroup;
                    tagGroup = unit.TagTableGroup;
                    break;
                }
            }
            if (blockGroup == null)
            {
                foreach (PlcSafetyUnit su in unitProvider.UnitGroup.SafetyUnits)
                {
                    if (su.Name == unitName)
                    {
                        blockGroup = su.BlockGroup;
                        extGroup = su.ExternalSourceGroup;
                        typeGroup = su.TypeGroup;
                        tagGroup = su.TagTableGroup;
                        break;
                    }
                }
            }
            if (blockGroup == null) { SendJson(resp, 404, new { error = "Unit not found: " + unitName }); return; }

            string exportDir = Path.Combine(tempDir, Guid.NewGuid().ToString());
            Directory.CreateDirectory(exportDir);
            string zipFile = Path.Combine(tempDir, MakeValidFileName(unitName) + ".zip");
            if (File.Exists(zipFile)) File.Delete(zipFile);

            Console.WriteLine("Exporting unit: " + unitName);
            try
            {
                using (var zip = ZipFile.Open(zipFile, ZipArchiveMode.Create))
                {
                    string prefix = MakeValidFileName(unitName) + "/";
                    ExportBlocksToZip(blockGroup, extGroup, "", prefix, zip, exportDir);
                    ExportTypesToZip(typeGroup, extGroup, prefix, zip, exportDir);
                    ExportTagTablesToZip(tagGroup.TagTables, prefix, zip, exportDir, tagGroup.Groups);
                }
                SendFile(resp, zipFile, MakeValidFileName(unitName) + ".zip");
            }
            finally
            {
                if (File.Exists(zipFile)) File.Delete(zipFile);
                if (Directory.Exists(exportDir)) Directory.Delete(exportDir, true);
            }
        }

        // ExportAllZip removed - use ExportAllXmlZip with root parameter instead

        static void SendFile(HttpListenerResponse resp, string filePath, string downloadName)
        {
            byte[] data = File.ReadAllBytes(filePath);
            resp.StatusCode = 200;
            if (downloadName.EndsWith(".zip"))
                resp.ContentType = "application/zip";
            else
                resp.ContentType = "application/octet-stream";
            resp.Headers.Add("Content-Disposition", "attachment; filename=\"" + downloadName + "\"");
            resp.ContentLength64 = data.Length;
            resp.OutputStream.Write(data, 0, data.Length);
            resp.Close();
        }

        static void SendFileWithExportReport(HttpListenerResponse resp, string filePath, string downloadName, int success, int failed)
        {
            byte[] data = File.ReadAllBytes(filePath);
            resp.StatusCode = 200;
            resp.ContentType = "application/zip";
            resp.Headers.Add("Content-Disposition", "attachment; filename=\"" + downloadName + "\"");
            resp.Headers.Add("X-Export-Success", success.ToString());
            resp.Headers.Add("X-Export-Failed", failed.ToString());
            resp.Headers.Add("Access-Control-Expose-Headers", "X-Export-Success, X-Export-Failed");
            resp.ContentLength64 = data.Length;
            resp.OutputStream.Write(data, 0, data.Length);
            resp.Close();
        }

        // ── HMI Routes ──

        static HmiTarget FindHmi(string deviceItemName)
        {
            EnsureConnected();
            foreach (Project proj in connectedPortal.Projects)
            {
                foreach (Device device in proj.Devices)
                {
                    foreach (DeviceItem di in device.DeviceItems)
                    {
                        if (di.Name == deviceItemName)
                        {
                            var sc = ((IEngineeringServiceProvider)di).GetService<SoftwareContainer>();
                            if (sc != null && sc.Software is HmiTarget hmi) return hmi;
                        }
                    }
                }
            }
            return null;
        }

        static void HandleHmiRoute(HttpListenerContext ctx, string route)
        {
            string hmiName = ctx.Request.QueryString["hmi"];
            if (string.IsNullOrEmpty(hmiName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?hmi= parameter" });
                return;
            }
            var hmi = FindHmi(hmiName);
            if (hmi == null)
            {
                SendJson(ctx.Response, 404, new { error = "HMI not found: " + hmiName });
                return;
            }

            if (route == "tags")
            {
                var tables = new List<object>();
                CollectHmiTags(hmi.TagFolder, tables);
                SendJson(ctx.Response, 200, new { count = tables.Count, tagTables = tables });
                return;
            }
            if (route == "textlists")
            {
                var lists = new List<object>();
                foreach (TextList tl in hmi.TextLists)
                    lists.Add(new { name = tl.Name });
                SendJson(ctx.Response, 200, new { count = lists.Count, textLists = lists });
                return;
            }

            SendJson(ctx.Response, 404, new { error = "Unknown HMI route: " + route });
        }

        static void CollectHmiTags(TagSystemFolder folder, List<object> result)
        {
            foreach (TagTable table in folder.TagTables)
                result.Add(new { name = table.Name });
            foreach (TagUserFolder uf in folder.Folders)
                CollectHmiTagsUser(uf, result);
        }

        static void CollectHmiTagsUser(TagUserFolder folder, List<object> result)
        {
            foreach (TagTable table in folder.TagTables)
                result.Add(new { name = table.Name });
            foreach (TagUserFolder uf in folder.Folders)
                CollectHmiTagsUser(uf, result);
        }

        // ── Hardware Catalog Search ──

        static void SearchCatalog(HttpListenerContext ctx)
        {
            string search = ctx.Request.QueryString["q"];
            if (string.IsNullOrEmpty(search))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?q= search term (e.g. q=1FK2)" });
                return;
            }
            EnsureConnected();

            var catalog = connectedPortal.HardwareCatalog;
            var entries = catalog.Find(search);
            var results = new List<object>();

            foreach (Siemens.Engineering.HW.HardwareCatalog.CatalogEntry entry in entries)
            {
                results.Add(new {
                    articleNumber = entry.ArticleNumber,
                    typeIdentifier = entry.TypeIdentifier,
                    typeName = entry.TypeName,
                    description = entry.Description,
                    version = entry.Version,
                    catalogPath = entry.CatalogPath
                });
            }

            SendJson(ctx.Response, 200, new { search, count = results.Count, entries = results });
        }

        static void SearchCompatible(HttpListenerContext ctx)
        {
            string search = ctx.Request.QueryString["q"] ?? "";
            string typeId = ctx.Request.QueryString["type"];
            if (string.IsNullOrEmpty(typeId))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?type= TypeIdentifier of the target device (e.g. type=OrderNumber:6SL3210-5HE10-4xFx/V5.2.3/S210). Use ?q= for search filter (optional)." });
                return;
            }
            EnsureConnected();

            var catalog = connectedPortal.HardwareCatalog;
            var entries = catalog.Find(search, typeId);
            var results = new List<object>();

            foreach (Siemens.Engineering.HW.HardwareCatalog.CatalogEntry entry in entries)
            {
                results.Add(new {
                    articleNumber = entry.ArticleNumber,
                    typeIdentifier = entry.TypeIdentifier,
                    typeName = entry.TypeName,
                    description = entry.Description,
                    version = entry.Version,
                    catalogPath = entry.CatalogPath
                });
            }

            SendJson(ctx.Response, 200, new { search, typeFilter = typeId, count = results.Count, entries = results });
        }

        // ── Device Item Exploration ──

        static void GetDeviceItems(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string projectName = ctx.Request.QueryString["project"];
            EnsureConnected();

            Device targetDevice = null;
            foreach (Project proj in connectedPortal.Projects)
            {
                if (projectName != null && proj.Name != projectName) continue;
                foreach (Device dev in proj.Devices)
                {
                    if (dev.Name == deviceName) { targetDevice = dev; break; }
                }
                if (targetDevice != null) break;
            }
            if (targetDevice == null) { SendJson(ctx.Response, 404, new { error = "Device not found: " + deviceName }); return; }

            var items = CollectDeviceItems(targetDevice.DeviceItems, 0);
            SendJson(ctx.Response, 200, new { device = targetDevice.Name, count = items.Count, items });
        }

        static List<object> CollectDeviceItems(DeviceItemComposition items, int depth)
        {
            var result = new List<object>();
            foreach (DeviceItem di in items)
            {
                var children = CollectDeviceItems(di.DeviceItems, depth + 1);
                result.Add(new {
                    name = di.Name,
                    type = di.TypeIdentifier,
                    classification = di.Classification.ToString(),
                    position = di.PositionNumber,
                    depth,
                    childCount = children.Count,
                    children
                });
            }
            return result;
        }

        static void GetDeviceItemAttributes(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string itemPath = ctx.Request.QueryString["item"];
            string projectName = ctx.Request.QueryString["project"];
            EnsureConnected();

            if (string.IsNullOrEmpty(deviceName) || string.IsNullOrEmpty(itemPath))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?device= and ?item= parameters. item is slash-separated path e.g. 'S210/PROFINET interface'" });
                return;
            }

            var di = FindDeviceItemByPath(deviceName, itemPath, projectName);
            if (di == null) { SendJson(ctx.Response, 404, new { error = "DeviceItem not found: " + itemPath }); return; }

            var attributes = ReadAllAttributes(di);
            SendJson(ctx.Response, 200, new { device = deviceName, item = itemPath, attributeCount = attributes.Count, attributes });
        }

        static void GetDeviceSubItems(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string itemPath = ctx.Request.QueryString["item"];
            string projectName = ctx.Request.QueryString["project"];
            int maxDepth = 2;
            string depthStr = ctx.Request.QueryString["depth"];
            if (depthStr != null) int.TryParse(depthStr, out maxDepth);

            EnsureConnected();

            if (string.IsNullOrEmpty(deviceName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?device= parameter" });
                return;
            }

            DeviceItemComposition startItems;
            string startName;

            if (string.IsNullOrEmpty(itemPath))
            {
                // Start from device root
                Device targetDevice = null;
                foreach (Project proj in connectedPortal.Projects)
                {
                    if (projectName != null && proj.Name != projectName) continue;
                    foreach (Device dev in proj.Devices)
                    {
                        if (dev.Name == deviceName) { targetDevice = dev; break; }
                    }
                    if (targetDevice != null) break;
                }
                if (targetDevice == null) { SendJson(ctx.Response, 404, new { error = "Device not found: " + deviceName }); return; }
                startItems = targetDevice.DeviceItems;
                startName = targetDevice.Name;
            }
            else
            {
                var di = FindDeviceItemByPath(deviceName, itemPath, projectName);
                if (di == null) { SendJson(ctx.Response, 404, new { error = "DeviceItem not found: " + itemPath }); return; }
                startItems = di.DeviceItems;
                startName = di.Name;
            }

            var items = CollectSubItemsWithAttributes(startItems, 0, maxDepth);
            SendJson(ctx.Response, 200, new { device = deviceName, root = startName, maxDepth, items });
        }

        static DeviceItem FindDeviceItemByPath(string deviceName, string itemPath, string projectName)
        {
            Device targetDevice = null;
            foreach (Project proj in connectedPortal.Projects)
            {
                if (projectName != null && proj.Name != projectName) continue;
                foreach (Device dev in proj.Devices)
                {
                    if (dev.Name == deviceName) { targetDevice = dev; break; }
                }
                if (targetDevice != null) break;
            }
            if (targetDevice == null) return null;

            string[] parts = itemPath.Split('/');
            DeviceItemComposition currentItems = targetDevice.DeviceItems;
            DeviceItem found = null;

            foreach (string part in parts)
            {
                found = null;
                foreach (DeviceItem di in currentItems)
                {
                    if (di.Name == part) { found = di; break; }
                }
                if (found == null) return null;
                currentItems = found.DeviceItems;
            }
            return found;
        }

        static List<object> ReadAllAttributes(DeviceItem di)
        {
            var attributes = new List<object>();
            try
            {
                var attrInfos = ((IEngineeringObject)di).GetAttributeInfos();
                foreach (var info in attrInfos)
                {
                    string val = null;
                    try
                    {
                        object raw = ((IEngineeringObject)di).GetAttribute(info.Name);
                        val = raw != null ? raw.ToString() : null;
                    }
                    catch { val = "<read error>"; }

                    attributes.Add(new { name = info.Name, value = val });
                }
            }
            catch (Exception ex)
            {
                attributes.Add(new { name = "_error", value = ex.Message });
            }
            return attributes;
        }

        static List<object> CollectSubItemsWithAttributes(DeviceItemComposition items, int depth, int maxDepth)
        {
            var result = new List<object>();
            foreach (DeviceItem di in items)
            {
                var attrs = ReadAllAttributes(di);
                var children = (depth < maxDepth) ? CollectSubItemsWithAttributes(di.DeviceItems, depth + 1, maxDepth) : new List<object>();

                result.Add(new {
                    name = di.Name,
                    type = di.TypeIdentifier,
                    classification = di.Classification.ToString(),
                    depth,
                    attributes = attrs,
                    children
                });
            }
            return result;
        }

        // ── Drive Configuration ──

        static DriveObjectContainer GetDriveContainer(string deviceName, string projectName = null)
        {
            foreach (Project proj in connectedPortal.Projects)
            {
                if (projectName != null && proj.Name != projectName) continue;
                foreach (Device dev in proj.Devices)
                {
                    if (dev.Name != deviceName) continue;
                    foreach (DeviceItem di in dev.DeviceItems)
                    {
                        var container = di.GetService<DriveObjectContainer>();
                        if (container != null) return container;
                        // Search one level deeper
                        foreach (DeviceItem child in di.DeviceItems)
                        {
                            container = child.GetService<DriveObjectContainer>();
                            if (container != null) return container;
                        }
                    }
                }
            }
            return null;
        }

        static void GetDriveObjects(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string projectName = ctx.Request.QueryString["project"];
            EnsureConnected();
            if (string.IsNullOrEmpty(deviceName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?device= parameter" });
                return;
            }

            var container = GetDriveContainer(deviceName, projectName);
            if (container == null)
            {
                SendJson(ctx.Response, 404, new { error = "No DriveObjectContainer found for device: " + deviceName });
                return;
            }

            var driveObjects = new List<object>();
            foreach (DriveObject dobj in container.DriveObjects)
            {
                var paramCount = 0;
                try { paramCount = dobj.Parameters.Count; } catch { }

                driveObjects.Add(new {
                    number = dobj.DriveObjectNumber,
                    parameterCount = paramCount
                });
            }

            SendJson(ctx.Response, 200, new { device = deviceName, driveObjectCount = driveObjects.Count, driveObjects });
        }

        static void GetDriveParameters(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string projectName = ctx.Request.QueryString["project"];
            string doNumStr = ctx.Request.QueryString["do"];
            string pNumStr = ctx.Request.QueryString["param"];
            string rangeStr = ctx.Request.QueryString["range"];
            EnsureConnected();

            if (string.IsNullOrEmpty(deviceName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?device= parameter. Optional: ?do=<driveObjectNumber>&param=<paramNumber>&range=300-320" });
                return;
            }

            var container = GetDriveContainer(deviceName, projectName);
            if (container == null)
            {
                SendJson(ctx.Response, 404, new { error = "No DriveObjectContainer found for device: " + deviceName });
                return;
            }

            int doNum = 0;
            if (!string.IsNullOrEmpty(doNumStr)) int.TryParse(doNumStr, out doNum);

            DriveObject targetDo = null;
            foreach (DriveObject dobj in container.DriveObjects)
            {
                if (dobj.DriveObjectNumber == doNum) { targetDo = dobj; break; }
            }
            if (targetDo == null)
            {
                // Fallback: first drive object
                foreach (DriveObject dobj in container.DriveObjects)
                {
                    targetDo = dobj;
                    break;
                }
            }
            if (targetDo == null)
            {
                SendJson(ctx.Response, 404, new { error = "No DriveObject found" });
                return;
            }

            // If specific param requested
            if (!string.IsNullOrEmpty(pNumStr))
            {
                int pNum;
                if (!int.TryParse(pNumStr, out pNum))
                {
                    SendJson(ctx.Response, 400, new { error = "Invalid param number" });
                    return;
                }
                var paramResult = ReadDriveParam(targetDo, pNum);
                SendJson(ctx.Response, 200, paramResult);
                return;
            }

            // If range requested (e.g. range=300-320)
            if (!string.IsNullOrEmpty(rangeStr))
            {
                var rangeParts = rangeStr.Split('-');
                int rangeStart, rangeEnd;
                if (rangeParts.Length == 2 && int.TryParse(rangeParts[0], out rangeStart) && int.TryParse(rangeParts[1], out rangeEnd))
                {
                    var rangeParams = new List<object>();
                    for (int i = rangeStart; i <= rangeEnd; i++)
                    {
                        var p = ReadDriveParam(targetDo, i);
                        if (p != null) rangeParams.Add(p);
                    }
                    SendJson(ctx.Response, 200, new { device = deviceName, driveObject = targetDo.DriveObjectNumber, range = rangeStr, count = rangeParams.Count, parameters = rangeParams });
                    return;
                }
            }

            // Default: return motor-related params (p300-p311)
            var motorParams = new List<object>();
            int[] defaultParams = { 300, 301, 302, 303, 304, 305, 306, 307, 308, 310, 311, 396, 397, 398 };
            foreach (int pn in defaultParams)
            {
                var p = ReadDriveParam(targetDo, pn);
                if (p != null) motorParams.Add(p);
            }

            SendJson(ctx.Response, 200, new { device = deviceName, driveObject = targetDo.DriveObjectNumber, count = motorParams.Count, parameters = motorParams });
        }

        static void DumpAllDriveParameters(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string projectName = ctx.Request.QueryString["project"];
            string doNumStr = ctx.Request.QueryString["do"];
            string filterStr = ctx.Request.QueryString["filter"]; // "enum" to show only params with EnumValueList, "motor" for motor-related
            EnsureConnected();

            if (string.IsNullOrEmpty(deviceName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?device=. Optional: ?do=1&filter=enum (show only params with enum lists)" });
                return;
            }

            var container = GetDriveContainer(deviceName, projectName);
            if (container == null) { SendJson(ctx.Response, 404, new { error = "No DriveObjectContainer: " + deviceName }); return; }

            int doNum = -1;
            if (!string.IsNullOrEmpty(doNumStr)) int.TryParse(doNumStr, out doNum);

            DriveObject targetDo = null;
            foreach (DriveObject dobj in container.DriveObjects)
            {
                if (doNum >= 0 && dobj.DriveObjectNumber == doNum) { targetDo = dobj; break; }
                if (targetDo == null) targetDo = dobj; // fallback to first
            }
            if (targetDo == null) { SendJson(ctx.Response, 404, new { error = "No DriveObject found" }); return; }

            bool filterEnum = (filterStr == "enum");
            bool filterMotor = (filterStr == "motor");

            var allParams = new List<object>();
            foreach (DriveParameter dp in targetDo.Parameters)
            {
                try
                {
                    var enumValues = new Dictionary<string, string>();
                    try
                    {
                        if (dp.EnumValueList != null)
                            foreach (var kvp in dp.EnumValueList)
                                enumValues[kvp.Key.ToString()] = kvp.Value;
                    }
                    catch { }

                    if (filterEnum && enumValues.Count == 0) continue;

                    string nameText = dp.Name ?? "";
                    string paramText = dp.ParameterText ?? "";
                    if (filterMotor)
                    {
                        string combined = (nameText + " " + paramText).ToLower();
                        if (!combined.Contains("motor") && !combined.Contains("encoder") && !combined.Contains("brake") &&
                            !combined.Contains("mlfb") && !combined.Contains("order") && !combined.Contains("article") &&
                            !combined.Contains("code") && !combined.Contains("selection") && !combined.Contains("type"))
                            continue;
                    }

                    object value = null;
                    try { value = dp.Value; } catch { value = "<err>"; }

                    allParams.Add(new {
                        number = dp.Number,
                        name = dp.Name,
                        text = dp.ParameterText,
                        value,
                        unit = dp.Unit,
                        enumCount = enumValues.Count,
                        enumValues = enumValues.Count > 0 ? (object)enumValues : null
                    });
                }
                catch { }
            }

            SendJson(ctx.Response, 200, new {
                device = deviceName,
                driveObject = targetDo.DriveObjectNumber,
                filter = filterStr ?? "none",
                totalParameterCount = targetDo.Parameters.Count,
                matchedCount = allParams.Count,
                parameters = allParams
            });
        }

        static object ReadDriveParam(DriveObject dobj, int paramNumber)
        {
            try
            {
                foreach (DriveParameter dp in dobj.Parameters)
                {
                    if (dp.Number == paramNumber)
                    {
                        object value = null;
                        try { value = dp.Value; } catch { value = "<read error>"; }

                        var enumValues = new Dictionary<string, string>();
                        try
                        {
                            if (dp.EnumValueList != null)
                            {
                                foreach (var kvp in dp.EnumValueList)
                                    enumValues[kvp.Key.ToString()] = kvp.Value;
                            }
                        }
                        catch { }

                        return new {
                            number = dp.Number,
                            name = dp.Name,
                            text = dp.ParameterText,
                            value,
                            unit = dp.Unit,
                            min = dp.MinValue,
                            max = dp.MaxValue,
                            arrayIndex = dp.ArrayIndex,
                            arrayLength = dp.ArrayLength,
                            enumValues = enumValues.Count > 0 ? (object)enumValues : null
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                return new { number = paramNumber, error = ex.Message };
            }
            return null;
        }

        static void GetMotorConfiguration(HttpListenerContext ctx)
        {
            string deviceName = ctx.Request.QueryString["device"];
            string projectName = ctx.Request.QueryString["project"];
            string doNumStr = ctx.Request.QueryString["do"];
            string ddsStr = ctx.Request.QueryString["dds"];
            EnsureConnected();

            if (string.IsNullOrEmpty(deviceName))
            {
                SendJson(ctx.Response, 400, new { error = "Missing ?device= parameter. Optional: ?do=<driveObjectNumber>&dds=<driveDataSetNumber>" });
                return;
            }

            var container = GetDriveContainer(deviceName, projectName);
            if (container == null)
            {
                SendJson(ctx.Response, 404, new { error = "No DriveObjectContainer found for device: " + deviceName });
                return;
            }

            int doNum = 0;
            if (!string.IsNullOrEmpty(doNumStr)) int.TryParse(doNumStr, out doNum);

            DriveObject targetDo = null;
            foreach (DriveObject dobj in container.DriveObjects)
            {
                if (dobj.DriveObjectNumber == doNum) { targetDo = dobj; break; }
            }
            if (targetDo == null)
            {
                foreach (DriveObject dobj in container.DriveObjects)
                { targetDo = dobj; break; }
            }
            if (targetDo == null)
            {
                SendJson(ctx.Response, 404, new { error = "No DriveObject found" });
                return;
            }

            ushort dds = 0;
            if (!string.IsNullOrEmpty(ddsStr)) ushort.TryParse(ddsStr, out dds);

            try
            {
                var dfi = targetDo.GetService<DriveFunctionInterface>();
                if (dfi == null)
                {
                    SendJson(ctx.Response, 404, new { error = "DriveFunctionInterface not available for this DriveObject" });
                    return;
                }

                // Debug: show what DFI properties are available
                var dfiInfo = new List<object>();
                try
                {
                    var dfiAttrs = ((IEngineeringObject)dfi).GetAttributeInfos();
                    foreach (var a in dfiAttrs)
                    {
                        string v = null;
                        try { var raw = ((IEngineeringObject)dfi).GetAttribute(a.Name); v = raw != null ? raw.ToString() : null; } catch { v = "<err>"; }
                        dfiInfo.Add(new { name = a.Name, value = v });
                    }
                }
                catch { }

                var dfiServices = new List<string>();
                try
                {
                    if (dfi.Commissioning != null) dfiServices.Add("Commissioning");
                } catch { }
                try
                {
                    if (dfi.DriveObjectFunctions != null) dfiServices.Add("DriveObjectFunctions");
                } catch { }
                try
                {
                    if (dfi.HardwareProjection != null) dfiServices.Add("HardwareProjection");
                } catch { }
                try
                {
                    if (dfi.FunctionInUse != null) dfiServices.Add("FunctionInUse");
                } catch { }
                try
                {
                    if (dfi.SafetyCommissioning != null) dfiServices.Add("SafetyCommissioning");
                } catch { }

                var hwProj = dfi.HardwareProjection;
                if (hwProj == null)
                {
                    SendJson(ctx.Response, 404, new {
                        error = "HardwareProjection not available",
                        driveObject = targetDo.DriveObjectNumber,
                        dfiAvailableServices = dfiServices,
                        dfiAttributes = dfiInfo
                    });
                    return;
                }

                var motorConfig = hwProj.GetCurrentMotorConfiguration(dds);
                if (motorConfig == null)
                {
                    SendJson(ctx.Response, 404, new { error = "No MotorConfiguration for DDS " + dds });
                    return;
                }

                var required = ReadConfigEntries(motorConfig.RequiredConfigurationEntries);
                var optional = ReadConfigEntries(motorConfig.OptionalConfigurationEntries);

                // Also get HardwareProjection attributes
                var hwAttrs = new List<object>();
                try
                {
                    var attrInfos = ((IEngineeringObject)hwProj).GetAttributeInfos();
                    foreach (var info in attrInfos)
                    {
                        string val = null;
                        try
                        {
                            var raw = ((IEngineeringObject)hwProj).GetAttribute(info.Name);
                            val = raw != null ? raw.ToString() : null;
                        }
                        catch { val = "<read error>"; }
                        hwAttrs.Add(new { name = info.Name, value = val });
                    }
                }
                catch { }

                SendJson(ctx.Response, 200, new {
                    device = deviceName,
                    driveObject = targetDo.DriveObjectNumber,
                    driveDataSet = dds,
                    requiredEntries = required,
                    optionalEntries = optional,
                    hwProjectionAttributes = hwAttrs
                });
            }
            catch (Exception ex)
            {
                SendJson(ctx.Response, 500, new { error = ex.Message, type = ex.GetType().Name });
            }
        }

        static List<object> ReadConfigEntries(ConfigurationEntryComposition entries)
        {
            var result = new List<object>();
            if (entries == null) return result;
            foreach (ConfigurationEntry entry in entries)
            {
                var enumValues = new Dictionary<string, string>();
                try
                {
                    if (entry.EnumValueList != null)
                    {
                        foreach (var kvp in entry.EnumValueList)
                            enumValues[kvp.Key.ToString()] = kvp.Value;
                    }
                }
                catch { }

                result.Add(new {
                    name = entry.Name,
                    number = entry.Number,
                    description = entry.Description,
                    value = entry.Value,
                    unit = entry.Unit,
                    min = entry.MinValue,
                    max = entry.MaxValue,
                    enumValues = enumValues.Count > 0 ? (object)enumValues : null
                });
            }
            return result;
        }

        // ── Helpers ──

        static void SendJson(HttpListenerResponse resp, int status, object data)
        {
            resp.StatusCode = status;
            resp.ContentType = "application/json; charset=utf-8";
            byte[] buf = Encoding.UTF8.GetBytes(json.Serialize(data));
            resp.ContentLength64 = buf.Length;
            resp.OutputStream.Write(buf, 0, buf.Length);
            resp.Close();
        }

        static void ServeOpenApiSpec(HttpListenerResponse resp)
        {
            string spec = @"{
  ""openapi"": ""3.0.3"",
  ""info"": {
    ""title"": ""TIA Portal Openness REST API"",
    ""version"": ""1.0.0"",
    ""description"": ""Siemens TIA Portal Openness (V20) 를 위한 REST API 브릿지입니다.\n\n실행 중인 TIA Portal 인스턴스에 연결하여 프로젝트 데이터를 조회하고, PLC 블록을 Import/Export 할 수 있습니다.\n\n---\n\n## 사전 요구사항\n\n| 항목 | 설명 |\n|---|---|\n| **TIA Portal V20** | 로컬 PC에 설치되어 있어야 합니다 |\n| **.NET Framework 4.8** | 런타임 필수 |\n| **Siemens TIA Openness** | Windows 그룹 'Siemens TIA Openness'에 사용자가 등록되어 있어야 합니다 |\n| **TiaApiServer.exe** | 로컬에서 실행 중이어야 합니다 (기본 포트: 8099) |\n\n## 주의사항\n\n- 이 API는 **로컬 PC에서만 동작**합니다. TIA Portal이 설치된 PC에서 TiaApiServer.exe를 실행한 뒤 localhost:8099로 접근해야 합니다.\n- 원격 서버에서는 사용할 수 없습니다. TIA Portal Openness API 자체가 로컬 프로세스 간 통신만 지원합니다.\n- API 서버를 먼저 실행한 후, /api/connect를 호출하여 TIA Portal 프로세스에 연결해야 다른 API를 사용할 수 있습니다.""
  },
  ""servers"": [{""url"": ""http://localhost:" + port + @""", ""description"": ""로컬 TIA API 서버""}],
  ""tags"": [
    {""name"": ""Connection"", ""description"": ""TIA Portal 프로세스 연결 관리""},
    {""name"": ""Setup"", ""description"": ""TIA Portal / 프로젝트 / 디바이스 신규 생성""},
    {""name"": ""Project"", ""description"": ""프로젝트 및 디바이스 정보 조회""},
    {""name"": ""Catalog"", ""description"": ""하드웨어 카탈로그 검색 및 호환성 확인""},
    {""name"": ""Device Explorer"", ""description"": ""디바이스 아이템 트리 및 속성 탐색""},
    {""name"": ""Drive"", ""description"": ""드라이브 오브젝트, 파라미터, 모터 설정 (Startdrive)""},
    {""name"": ""PLC"", ""description"": ""PLC 블록, 데이터 타입, 태그, Software Unit 조회""},
    {""name"": ""Export"", ""description"": ""PLC 프로그램 내보내기 (SCL/DB/STL/XML ZIP)""},
    {""name"": ""Import"", ""description"": ""PLC 프로그램 가져오기 (XML/ZIP)""},
    {""name"": ""HMI"", ""description"": ""HMI 태그 및 텍스트 리스트""},
    {""name"": ""DLL"", ""description"": ""Siemens Engineering DLL 관리 - DLL 상태 확인 및 업로드""}
  ],
  ""paths"": {
    ""/api/dll/status"": {
      ""get"": {
        ""tags"": [""DLL""],
        ""summary"": ""DLL 상태 확인"",
        ""description"": ""Siemens.Engineering.dll 등 필수 DLL의 존재 여부를 확인합니다. TIA Portal이 설치되어 있으면 자동으로 감지되며, 설치되어 있지 않은 경우 /api/dll/upload로 수동 업로드해야 합니다."",
        ""responses"": {""200"": {""description"": ""DLL 상태 정보 - 각 DLL의 존재 여부와 경로""}}
      }
    },
    ""/api/dll/upload"": {
      ""post"": {
        ""tags"": [""DLL""],
        ""summary"": ""DLL 파일 업로드"",
        ""description"": ""Siemens.Engineering.dll 파일을 업로드합니다. TIA Portal이 설치된 PC의 다음 경로에서 DLL을 찾을 수 있습니다:\\n\\n- C:\\\\Program Files\\\\Siemens\\\\Automation\\\\Portal V20\\\\PublicAPI\\\\V20\\\\Siemens.Engineering.dll\\n\\n업로드된 DLL은 서버의 dll/ 폴더에 저장됩니다."",
        ""parameters"": [
          {""name"": ""name"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""DLL 파일명 (예: Siemens.Engineering.dll)""}
        ],
        ""requestBody"": {
          ""required"": true,
          ""content"": {
            ""application/octet-stream"": {
              ""schema"": {""type"": ""string"", ""format"": ""binary""}
            }
          }
        },
        ""responses"": {
          ""200"": {""description"": ""업로드 성공 - 파일명, 크기, 저장 경로 반환""},
          ""400"": {""description"": ""파일명 누락 또는 허용되지 않는 DLL""}
        }
      }
    },
    ""/api/processes"": {
      ""get"": {
        ""tags"": [""Connection""],
        ""summary"": ""실행 중인 TIA Portal 프로세스 목록 조회"",
        ""description"": ""현재 PC에서 실행 중인 모든 TIA Portal 프로세스를 검색합니다. 각 프로세스의 ID와 열려있는 프로젝트 경로를 반환합니다."",
        ""responses"": {""200"": {""description"": ""TIA Portal 프로세스 목록""}}
      }
    },
    ""/api/connect"": {
      ""get"": {
        ""tags"": [""Connection""],
        ""summary"": ""TIA Portal 프로세스에 연결 (Attach)"",
        ""description"": ""지정된 TIA Portal 프로세스에 연결합니다. processId를 생략하면 첫 번째 프로세스에 자동 연결됩니다. 이 API를 호출해야 다른 모든 API를 사용할 수 있습니다."",
        ""parameters"": [{""name"": ""processId"", ""in"": ""query"", ""schema"": {""type"": ""integer""}, ""description"": ""연결할 TIA Portal 프로세스 ID (생략 시 첫 번째 프로세스에 자동 연결)""}],
        ""responses"": {""200"": {""description"": ""연결 성공""}, ""404"": {""description"": ""실행 중인 TIA Portal을 찾을 수 없음""}}
      }
    },
    ""/api/status"": {
      ""get"": {
        ""tags"": [""Connection""],
        ""summary"": ""현재 연결 상태 확인"",
        ""description"": ""API 서버가 TIA Portal에 연결되어 있는지, 열려있는 프로젝트가 몇 개인지 확인합니다."",
        ""responses"": {""200"": {""description"": ""연결 상태 정보""}}
      }
    },
    ""/api/projects"": {
      ""get"": {
        ""tags"": [""Project""],
        ""summary"": ""열려있는 프로젝트 목록 조회"",
        ""description"": ""현재 TIA Portal에서 열려있는 모든 프로젝트의 이름과 경로를 반환합니다."",
        ""responses"": {""200"": {""description"": ""프로젝트 목록""}}
      }
    },
    ""/api/devices"": {
      ""get"": {
        ""tags"": [""Project""],
        ""summary"": ""프로젝트의 모든 디바이스 목록 조회"",
        ""description"": ""프로젝트에 포함된 모든 디바이스(PLC, HMI, 드라이브 등)와 하위 아이템을 조회합니다. 반환된 PLC/HMI 아이템 이름을 다른 API의 plc 또는 hmi 파라미터에 사용하세요."",
        ""parameters"": [{""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""특정 프로젝트만 필터링 (생략 시 모든 프로젝트)""}],
        ""responses"": {""200"": {""description"": ""디바이스 목록 및 하위 아이템 정보""}}
      }
    },
    ""/api/portal/new"": {
      ""get"": {
        ""tags"": [""Setup""],
        ""summary"": ""새 TIA Portal 인스턴스 시작"",
        ""description"": ""새로운 TIA Portal 프로세스를 실행하고 자동으로 연결합니다. 기존 연결은 새 연결로 대체됩니다. WithUserInterface: UI 표시 (기본값), WithoutUserInterface: 백그라운드 실행 (자동화에 적합). TIA Portal 시작에 수십 초가 소요될 수 있습니다."",
        ""parameters"": [{""name"": ""mode"", ""in"": ""query"", ""schema"": {""type"": ""string"", ""enum"": [""WithUserInterface"", ""WithoutUserInterface""], ""default"": ""WithUserInterface""}, ""description"": ""실행 모드 - WithUserInterface(UI 표시) 또는 WithoutUserInterface(백그라운드)""}],
        ""responses"": {""200"": {""description"": ""새 TIA Portal이 시작되고 연결됨""}}
      }
    },
    ""/api/project/create"": {
      ""get"": {
        ""tags"": [""Setup""],
        ""summary"": ""새 프로젝트 생성"",
        ""description"": ""TIA Portal에 새 프로젝트를 생성합니다. 지정된 경로에 프로젝트 폴더가 생성됩니다."",
        ""parameters"": [
          {""name"": ""path"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""프로젝트를 저장할 디렉토리 경로 (예: C:\\\\Projects)""},
          {""name"": ""name"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""생성할 프로젝트 이름""}
        ],
        ""responses"": {""200"": {""description"": ""프로젝트 생성 성공""}, ""400"": {""description"": ""path 또는 name 파라미터 누락""}}
      }
    },
    ""/api/project/add-device"": {
      ""get"": {
        ""tags"": [""Setup""],
        ""summary"": ""프로젝트에 디바이스(PLC) 추가"",
        ""description"": ""프로젝트에 새 디바이스를 추가합니다. 기본값은 S7-1500 PLC입니다. Order Number 예시: S7-1517: OrderNumber:6ES7 517-3AP00-0AB0/V3.1"",
        ""parameters"": [
          {""name"": ""name"", ""in"": ""query"", ""schema"": {""type"": ""string"", ""default"": ""NewDevice""}, ""description"": ""디바이스 이름""},
          {""name"": ""type"", ""in"": ""query"", ""schema"": {""type"": ""string"", ""default"": ""System:Device.S71500""}, ""description"": ""디바이스 타입 식별자""},
          {""name"": ""order"", ""in"": ""query"", ""schema"": {""type"": ""string"", ""default"": ""OrderNumber:6ES7 517-3AP00-0AB0/V3.1""}, ""description"": ""주문번호 및 펌웨어 버전""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 첫 번째 프로젝트)""}
        ],
        ""responses"": {""200"": {""description"": ""디바이스 추가 성공""}, ""404"": {""description"": ""프로젝트를 찾을 수 없음""}}
      }
    },
    ""/api/catalog/search"": {
      ""get"": {
        ""tags"": [""Catalog""],
        ""summary"": ""하드웨어 카탈로그 검색"",
        ""description"": ""TIA Portal 하드웨어 카탈로그에서 키워드로 검색합니다. 예: q=1FK2 (모터), q=S210 (드라이브)"",
        ""parameters"": [
          {""name"": ""q"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""검색어 (예: 1FK2, S210, 6ES7)""}
        ],
        ""responses"": {""200"": {""description"": ""카탈로그 항목 목록 (articleNumber, typeName, description, version 포함)""}}
      }
    },
    ""/api/catalog/compatible"": {
      ""get"": {
        ""tags"": [""Catalog""],
        ""summary"": ""특정 디바이스에 호환되는 모듈 검색"",
        ""description"": ""특정 디바이스의 TypeIdentifier를 기준으로 호환 가능한 모듈을 검색합니다. S210에 꽂을 수 있는 1FK2 모터 목록 등을 조회할 수 있습니다."",
        ""parameters"": [
          {""name"": ""q"", ""in"": ""query"", ""schema"": {""type"": ""string"", ""default"": """"}, ""description"": ""검색 필터 (빈 값이면 전체 호환 목록)""},
          {""name"": ""type"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""대상 디바이스의 TypeIdentifier (예: OrderNumber:6SL3210-5HE10-4xFx/V5.2.3/S210)""}
        ],
        ""responses"": {""200"": {""description"": ""호환 가능한 카탈로그 항목 목록""}}
      }
    },
    ""/api/device/items"": {
      ""get"": {
        ""tags"": [""Device Explorer""],
        ""summary"": ""디바이스의 전체 아이템 트리 조회"",
        ""description"": ""디바이스 내부의 모든 DeviceItem(모듈, 서브모듈, 인터페이스, 포트)을 재귀적으로 탐색합니다. SINAMICS S210 같은 드라이브의 내부 구조를 파악할 때 유용합니다."",
        ""parameters"": [
          {""name"": ""device"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""디바이스 이름 (예: S210_1)""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 전체 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""DeviceItem 트리 - name, type, classification, position, children 포함""}}
      }
    },
    ""/api/device/item-attributes"": {
      ""get"": {
        ""tags"": [""Device Explorer""],
        ""summary"": ""특정 DeviceItem의 모든 속성 읽기"",
        ""description"": ""DeviceItem의 읽기 가능한 모든 속성(이름 + 값)을 반환합니다. 슬래시(/)로 구분된 경로를 사용하여 하위 아이템을 지정합니다. 예: device=S210_1&item=S210/PROFINET interface"",
        ""parameters"": [
          {""name"": ""device"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""디바이스 이름""},
          {""name"": ""item"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""슬래시(/)로 구분된 DeviceItem 경로 (예: S210/PROFINET interface)""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 전체 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""속성 이름/값 쌍 목록""}}
      }
    },
    ""/api/device/subitems"": {
      ""get"": {
        ""tags"": [""Device Explorer""],
        ""summary"": ""하위 아이템 + 전체 속성 덤프 (Deep Dump)"",
        ""description"": ""DeviceItem을 재귀적으로 탐색하면서 각 아이템의 모든 속성을 함께 반환합니다. 모터 선택 데이터, 드라이브 파라미터 등을 발견할 때 유용합니다. depth 파라미터로 탐색 깊이를 제어하세요."",
        ""parameters"": [
          {""name"": ""device"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""디바이스 이름""},
          {""name"": ""item"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""시작 DeviceItem 경로 (생략 시 디바이스 루트부터 시작)""},
          {""name"": ""depth"", ""in"": ""query"", ""schema"": {""type"": ""integer"", ""default"": 2}, ""description"": ""최대 재귀 깊이 (기본값: 2)""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 전체 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""전체 속성이 포함된 DeviceItem 트리""}}
      }
    },
    ""/api/drive/objects"": {
      ""get"": {
        ""tags"": [""Drive""],
        ""summary"": ""디바이스의 DriveObject 목록 조회"",
        ""description"": ""지정된 디바이스에서 Startdrive를 통해 접근 가능한 모든 DriveObject를 나열합니다."",
        ""parameters"": [
          {""name"": ""device"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""디바이스 이름 (예: S210_1)""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 전체 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""DriveObject 목록 (파라미터 개수 포함)""}}
      }
    },
    ""/api/drive/parameters"": {
      ""get"": {
        ""tags"": [""Drive""],
        ""summary"": ""드라이브 파라미터 읽기"",
        ""description"": ""DriveObject에서 드라이브 파라미터를 읽습니다. 기본값: 모터 관련 파라미터 (p300~p311). 특정 파라미터: param=300. 범위 조회: range=300-320"",
        ""parameters"": [
          {""name"": ""device"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""디바이스 이름""},
          {""name"": ""do"", ""in"": ""query"", ""schema"": {""type"": ""integer""}, ""description"": ""DriveObject 번호 (기본값: 0)""},
          {""name"": ""param"", ""in"": ""query"", ""schema"": {""type"": ""integer""}, ""description"": ""특정 파라미터 번호 (예: 300)""},
          {""name"": ""range"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""파라미터 범위 (예: 300-320)""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 전체 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""드라이브 파라미터 값, 단위, Enum 목록""}}
      }
    },
    ""/api/drive/motor-config"": {
      ""get"": {
        ""tags"": [""Drive""],
        ""summary"": ""현재 모터 설정 조회"",
        ""description"": ""HardwareProjection을 통해 현재 모터 설정을 조회합니다. 필수(Required) 및 선택(Optional) ConfigurationEntry와 각 항목의 Enum 값 목록을 반환합니다."",
        ""parameters"": [
          {""name"": ""device"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""디바이스 이름""},
          {""name"": ""do"", ""in"": ""query"", ""schema"": {""type"": ""integer""}, ""description"": ""DriveObject 번호 (기본값: 0)""},
          {""name"": ""dds"", ""in"": ""query"", ""schema"": {""type"": ""integer""}, ""description"": ""Drive Data Set 번호 (기본값: 0)""},
          {""name"": ""project"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""프로젝트 이름 (생략 시 전체 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""모터 설정 - 필수/선택 항목과 Enum 값 목록""}}
      }
    },
    ""/api/plc/create-unit"": {
      ""get"": {
        ""tags"": [""Import""],
        ""summary"": ""새 Software Unit 생성"",
        ""description"": ""PLC에 새로운 Software Unit을 생성합니다. Software Unit은 PLC 프로그램을 기능별로 모듈화하여 독립적으로 관리할 수 있게 해줍니다."",
        ""parameters"": [
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""},
          {""name"": ""name"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""생성할 Software Unit 이름""}
        ],
        ""responses"": {""200"": {""description"": ""Unit 생성 성공""}}
      }
    },
    ""/api/plc/blocks"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""PLC 블록 전체 목록 조회 (재귀)"",
        ""description"": ""PLC의 모든 블록(OB, FB, FC, DB)을 그룹 구조를 포함하여 재귀적으로 조회합니다. 각 블록의 이름, 번호, 프로그래밍 언어, 타입, 그룹 경로를 반환합니다."",
        ""parameters"": [{""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름 (예: PLC_1, SYNEX_Sample_PLC)""}],
        ""responses"": {""200"": {""description"": ""블록 목록 - name, number, language, type, group 포함""}}
      }
    },
    ""/api/plc/block-detail"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""특정 블록 상세 정보 조회"",
        ""description"": ""특정 PLC 블록의 상세 정보를 조회합니다. 블록 이름, 번호, 프로그래밍 언어, 타입, 수정일시, 인터페이스(Input/Output/InOut/Static/Temp) 등을 반환합니다."",
        ""parameters"": [
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""},
          {""name"": ""name"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""블록 이름 (예: Main, FB_Motor)""}
        ],
        ""responses"": {""200"": {""description"": ""블록 상세 정보""}}
      }
    },
    ""/api/plc/types"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""PLC 데이터 타입(UDT) 전체 목록 조회"",
        ""description"": ""PLC에 정의된 모든 사용자 정의 데이터 타입(UDT)을 조회합니다."",
        ""parameters"": [{""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}],
        ""responses"": {""200"": {""description"": ""UDT 목록""}}
      }
    },
    ""/api/plc/tags"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""PLC 태그 테이블 전체 목록 조회"",
        ""description"": ""PLC의 모든 태그 테이블과 각 테이블에 포함된 태그(이름, 데이터 타입, 주소, 코멘트)를 조회합니다. 하위 그룹의 태그 테이블도 재귀적으로 포함됩니다."",
        ""parameters"": [{""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}],
        ""responses"": {""200"": {""description"": ""태그 테이블 목록 (태그 상세 포함)""}}
      }
    },
    ""/api/plc/units"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""Software Unit 전체 목록 조회"",
        ""description"": ""PLC에 생성된 모든 Software Unit을 조회합니다. Software Unit은 PLC 프로그램을 모듈화하여 관리하는 단위입니다."",
        ""parameters"": [{""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}],
        ""responses"": {""200"": {""description"": ""Software Unit 목록 (블록/타입/태그 개수 포함)""}}
      }
    },
    ""/api/plc/units/{unitName}/blocks"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""특정 Software Unit의 블록 목록 조회"",
        ""description"": ""지정된 Software Unit에 포함된 모든 블록을 조회합니다."",
        ""parameters"": [
          {""name"": ""unitName"", ""in"": ""path"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름""},
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}
        ],
        ""responses"": {""200"": {""description"": ""해당 Unit의 블록 목록""}}
      }
    },
    ""/api/plc/units/{unitName}/types"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""특정 Software Unit의 데이터 타입(UDT) 목록 조회"",
        ""description"": ""지정된 Software Unit에 정의된 모든 사용자 정의 데이터 타입을 조회합니다."",
        ""parameters"": [
          {""name"": ""unitName"", ""in"": ""path"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름""},
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}
        ],
        ""responses"": {""200"": {""description"": ""해당 Unit의 UDT 목록""}}
      }
    },
    ""/api/plc/units/{unitName}/tags"": {
      ""get"": {
        ""tags"": [""PLC""],
        ""summary"": ""특정 Software Unit의 태그 테이블 목록 조회"",
        ""description"": ""지정된 Software Unit에 포함된 모든 태그 테이블과 태그를 조회합니다."",
        ""parameters"": [
          {""name"": ""unitName"", ""in"": ""path"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름""},
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}
        ],
        ""responses"": {""200"": {""description"": ""해당 Unit의 태그 테이블 목록""}}
      }
    },
    ""/api/plc/export/block"": {
      ""get"": {
        ""tags"": [""Export""],
        ""summary"": ""단일 블록을 소스 파일로 다운로드"",
        ""description"": ""특정 PLC 블록을 소스 코드 파일로 다운로드합니다. SCL 블록은 .scl, DB 블록은 .db, STL 블록은 .awl 형식으로 내보냅니다. 주의: LAD/FBD 블록은 이 API로 내보낼 수 없습니다. /api/plc/export/xml을 사용하세요."",
        ""parameters"": [
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""},
          {""name"": ""name"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""블록 이름 (예: Main, FB_Motor)""},
          {""name"": ""unit"", ""in"": ""query"", ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름 (생략 시 루트 블록에서 검색)""}
        ],
        ""responses"": {""200"": {""description"": ""소스 파일 다운로드 (.scl / .db / .awl)""}}
      }
    },
    ""/api/plc/units/{unitName}/export"": {
      ""get"": {
        ""tags"": [""Export""],
        ""summary"": ""Software Unit 전체를 ZIP으로 다운로드 (소스 파일)"",
        ""description"": ""Software Unit의 모든 블록, 타입, 태그 테이블을 ZIP 파일로 다운로드합니다. ZIP 내부 폴더: scl/, db/, stl/, udt/, tag_tables/. 주의: LAD/FBD 블록은 포함되지 않습니다. 전체 블록이 필요하면 /export-xml을 사용하세요."",
        ""parameters"": [
          {""name"": ""unitName"", ""in"": ""path"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름""},
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}
        ],
        ""responses"": {""200"": {""description"": ""ZIP 파일 다운로드""}}
      }
    },
    ""/api/plc/export/xml"": {
      ""get"": {
        ""tags"": [""Export""],
        ""summary"": ""PLC 전체를 XML ZIP으로 다운로드 (LAD/FBD 포함)"",
        ""description"": ""PLC의 모든 블록을 SimaticML XML 형식으로 내보냅니다. LAD/FBD 블록을 포함한 모든 블록 타입을 지원합니다. Export 범위: all(전체, 기본값), program_blocks(루트만), software_units(유닛만). 이 ZIP을 /api/plc/import/xml로 다시 Import할 수 있습니다."",
        ""parameters"": [
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""},
          {""name"": ""root"", ""in"": ""query"", ""schema"": {""type"": ""string"", ""enum"": [""all"", ""program_blocks"", ""software_units""], ""default"": ""all""}, ""description"": ""Export 범위: all(전체), program_blocks(루트만), software_units(유닛만)""}
        ],
        ""responses"": {""200"": {""description"": ""ZIP 파일 다운로드 (XML 형식)""}}
      }
    },
    ""/api/plc/units/{unitName}/export-xml"": {
      ""get"": {
        ""tags"": [""Export""],
        ""summary"": ""특정 Software Unit을 XML ZIP으로 다운로드 (LAD/FBD 포함)"",
        ""description"": ""특정 Software Unit의 모든 블록, 타입, 태그를 SimaticML XML 형식의 ZIP으로 내보냅니다. LAD/FBD를 포함한 모든 블록 타입을 지원합니다."",
        ""parameters"": [
          {""name"": ""unitName"", ""in"": ""path"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름""},
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""}
        ],
        ""responses"": {""200"": {""description"": ""ZIP 파일 다운로드 (XML 형식)""}}
      }
    },
    ""/api/plc/import/xml"": {
      ""get"": {
        ""tags"": [""Import""],
        ""summary"": ""로컬 XML 파일 또는 ZIP을 Import"",
        ""description"": ""로컬 파일 시스템의 XML 또는 ZIP을 PLC에 Import합니다. 단일 XML: 블록/타입/태그를 자동 감지. ZIP: /export/xml 구조를 인식하며 타입->태그->블록 순서로 의존성을 고려하여 처리. Software Unit과 블록 그룹은 자동 생성됩니다."",
        ""parameters"": [
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""},
          {""name"": ""file"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""로컬 파일 경로 (.xml 또는 .zip)""}
        ],
        ""responses"": {""200"": {""description"": ""Import 결과 - 파일별 성공/실패 상태""}, ""400"": {""description"": ""파일 경로 누락 또는 잘못된 확장자""}, ""500"": {""description"": ""TIA Portal Import 오류""}}
      }
    },
    ""/api/plc/import/upload"": {
      ""post"": {
        ""tags"": [""Import""],
        ""summary"": ""XML 파일 업로드를 통한 Import"",
        ""description"": ""HTTP 요청 본문에 XML 내용을 담아 직접 업로드합니다. target으로 Import 대상을 지정: blocks(블록), tags(태그), types(UDT). unit을 지정하면 해당 Software Unit에, 생략하면 PLC 루트에 Import됩니다."",
        ""parameters"": [
          {""name"": ""plc"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""PLC 디바이스 아이템 이름""},
          {""name"": ""target"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string"", ""enum"": [""blocks"", ""tags"", ""types""]}, ""description"": ""Import 대상: blocks(블록), tags(태그 테이블), types(UDT)""},
          {""name"": ""unit"", ""in"": ""query"", ""required"": false, ""schema"": {""type"": ""string""}, ""description"": ""Software Unit 이름 (생략 시 PLC 루트에 Import)""}
        ],
        ""requestBody"": {
          ""required"": true,
          ""description"": ""SimaticML XML 형식의 파일 내용"",
          ""content"": {
            ""application/xml"": {
              ""schema"": {""type"": ""string"", ""format"": ""binary""},
              ""example"": ""<?xml version='1.0' encoding='utf-8'?><Document>...</Document>""
            }
          }
        },
        ""responses"": {
          ""200"": {""description"": ""Import 성공 - 블록 종류와 Unit 정보 반환""},
          ""400"": {""description"": ""요청 본문이 비어있거나 필수 파라미터 누락""},
          ""404"": {""description"": ""지정된 Software Unit을 찾을 수 없음""},
          ""500"": {""description"": ""TIA Portal Import 오류""}
        }
      }
    },
    ""/api/hmi/tags"": {
      ""get"": {
        ""tags"": [""HMI""],
        ""summary"": ""HMI 태그 테이블 목록 조회"",
        ""description"": ""HMI 디바이스의 모든 태그 테이블을 조회합니다. 시스템 폴더와 사용자 폴더의 태그 테이블을 재귀적으로 탐색합니다."",
        ""parameters"": [{""name"": ""hmi"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""HMI 디바이스 아이템 이름 (예: HMI_1)""}],
        ""responses"": {""200"": {""description"": ""HMI 태그 테이블 목록""}}
      }
    },
    ""/api/hmi/textlists"": {
      ""get"": {
        ""tags"": [""HMI""],
        ""summary"": ""HMI 텍스트 리스트 목록 조회"",
        ""description"": ""HMI 디바이스에 정의된 모든 텍스트 리스트를 조회합니다. 텍스트 리스트는 HMI 화면에서 상태값을 사용자 친화적인 텍스트로 표시할 때 사용됩니다."",
        ""parameters"": [{""name"": ""hmi"", ""in"": ""query"", ""required"": true, ""schema"": {""type"": ""string""}, ""description"": ""HMI 디바이스 아이템 이름 (예: HMI_1)""}],
        ""responses"": {""200"": {""description"": ""HMI 텍스트 리스트 목록""}}
      }
    }
  }
}";
            resp.StatusCode = 200;
            resp.ContentType = "application/json; charset=utf-8";
            byte[] buf = Encoding.UTF8.GetBytes(spec);
            resp.ContentLength64 = buf.Length;
            resp.OutputStream.Write(buf, 0, buf.Length);
            resp.Close();
        }

        static void ServeSwaggerUI(HttpListenerResponse resp)
        {
            string html = @"<!DOCTYPE html>
<html lang=""ko"">
<head>
<meta charset=""UTF-8"">
<title>TIA Portal Openness REST API</title>
<link rel=""stylesheet"" href=""https://unpkg.com/swagger-ui-dist@5/swagger-ui.css"" />
<style>
body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
.topbar-wrapper img { content: url('data:image/svg+xml;utf8,<svg xmlns=""http://www.w3.org/2000/svg"" width=""40"" height=""40""><rect width=""40"" height=""40"" rx=""6"" fill=""%2349cc90""/><text x=""20"" y=""27"" text-anchor=""middle"" fill=""white"" font-size=""20"" font-weight=""bold"">T</text></svg>'); }
.swagger-ui .topbar { background: #1b1b1b; }
.notice-banner { background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); border-bottom: 3px solid #f0ad4e; padding: 16px 24px; display: flex; align-items: flex-start; gap: 12px; }
.notice-banner .icon { font-size: 24px; flex-shrink: 0; margin-top: 2px; }
.notice-banner .content { flex: 1; }
.notice-banner h3 { margin: 0 0 8px 0; color: #856404; font-size: 15px; }
.notice-banner ul { margin: 6px 0 0 0; padding-left: 18px; color: #664d03; font-size: 13px; line-height: 1.8; }
.dll-panel { background: #f8f9fa; border-bottom: 2px solid #dee2e6; padding: 20px 24px; }
.dll-panel h3 { margin: 0 0 12px 0; font-size: 15px; color: #333; }
.dll-panel .status-area { margin-bottom: 12px; font-size: 13px; color: #555; }
.dll-panel .status-item { padding: 4px 0; }
.dll-panel .found { color: #28a745; }
.dll-panel .missing { color: #dc3545; }
.dll-panel .upload-area { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.dll-panel input[type=file] { font-size: 13px; }
.dll-panel button { background: #0d6efd; color: white; border: none; padding: 8px 20px; border-radius: 4px; cursor: pointer; font-size: 13px; }
.dll-panel button:hover { background: #0b5ed7; }
.dll-panel button:disabled { background: #6c757d; cursor: not-allowed; }
.dll-panel .upload-result { margin-top: 8px; font-size: 13px; padding: 8px 12px; border-radius: 4px; }
</style>
</head>
<body>
<div class=""notice-banner"">
  <span class=""icon"">&#9888;&#65039;</span>
  <div class=""content"">
    <h3>&#47196;&#52972; &#54872;&#44221; &#51204;&#50857; API</h3>
    <ul>
      <li>Siemens TIA Portal V20&#51060; &#47196;&#52972; PC&#50640; &#49444;&#52824;&#46104;&#50612; &#51080;&#50612;&#50556; &#54633;&#45768;&#45796;</li>
      <li>Windows &#49324;&#50857;&#51088;&#44032; 'Siemens TIA Openness' &#44536;&#47353;&#50640; &#46321;&#47197;&#46104;&#50612; &#51080;&#50612;&#50556; &#54633;&#45768;&#45796;</li>
      <li>TiaApiServer.exe&#47484; &#49892;&#54665;&#54620; &#54980; /api/connect&#47484; &#54840;&#52636;&#54616;&#50668; TIA Portal&#50640; &#50672;&#44208;&#54616;&#49464;&#50836;</li>
    </ul>
  </div>
</div>
<div class=""dll-panel"">
  <h3>Siemens Engineering DLL</h3>
  <div class=""status-area"" id=""dll-status"">DLL &#49345;&#53468; &#54869;&#51064; &#51473;...</div>
  <div class=""upload-area"">
    <input type=""file"" id=""dll-file"" accept="".dll"" />
    <button id=""dll-upload-btn"" onclick=""uploadDll()"">&#50629;&#47196;&#46300;</button>
  </div>
  <div id=""dll-upload-result""></div>
</div>
<div id=""swagger-ui""></div>
<script src=""https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js""></script>
<script>
SwaggerUIBundle({
  url: '/api/openapi.json',
  dom_id: '#swagger-ui',
  deepLinking: true,
  presets: [SwaggerUIBundle.presets.apis, SwaggerUIBundle.SwaggerUIStandalonePreset],
  layout: 'BaseLayout',
  defaultModelsExpandDepth: -1,
  docExpansion: 'list',
  tryItOutEnabled: true
});

function checkDllStatus() {
  fetch('/api/dll/status').then(r => r.json()).then(data => {
    var html = '';
    data.dlls.forEach(function(d) {
      if (d.found) html += '<div class=""status-item found"">&#10003; ' + d.name + ' (' + d.location + ')</div>';
      else html += '<div class=""status-item missing"">&#10007; ' + d.name + ' - &#50629;&#47196;&#46300; &#54596;&#50836;</div>';
    });
    if (data.ready) html = '<div class=""status-item found""><strong>&#10003; &#47784;&#46304; DLL&#51060; &#51456;&#48708;&#46104;&#50632;&#49845;&#45768;&#45796;</strong></div>' + html;
    document.getElementById('dll-status').innerHTML = html;
  }).catch(function() {
    document.getElementById('dll-status').innerHTML = '<div class=""status-item missing"">API &#49436;&#48260;&#50640; &#50672;&#44208;&#54624; &#49688; &#50630;&#49845;&#45768;&#45796;</div>';
  });
}

function uploadDll() {
  var fileInput = document.getElementById('dll-file');
  var resultDiv = document.getElementById('dll-upload-result');
  if (!fileInput.files.length) { resultDiv.innerHTML = '<div style=""color:#dc3545"">&#54028;&#51068;&#51012; &#49440;&#53469;&#54644;&#51452;&#49464;&#50836;</div>'; return; }
  var file = fileInput.files[0];
  if (!file.name.endsWith('.dll')) { resultDiv.innerHTML = '<div style=""color:#dc3545"">.dll &#54028;&#51068;&#47564; &#50629;&#47196;&#46300; &#44032;&#45733;&#54633;&#45768;&#45796;</div>'; return; }
  var btn = document.getElementById('dll-upload-btn');
  btn.disabled = true; btn.textContent = '&#50629;&#47196;&#46300; &#51473;...';
  fetch('/api/dll/upload?name=' + encodeURIComponent(file.name), { method: 'POST', body: file })
    .then(r => r.json()).then(data => {
      if (data.error) resultDiv.innerHTML = '<div class=""upload-result"" style=""background:#f8d7da;color:#842029"">' + data.error + '</div>';
      else { resultDiv.innerHTML = '<div class=""upload-result"" style=""background:#d1e7dd;color:#0f5132"">' + data.message + ' (' + data.size + ' bytes)</div>'; checkDllStatus(); }
      btn.disabled = false; btn.textContent = '&#50629;&#47196;&#46300;';
    }).catch(function(e) {
      resultDiv.innerHTML = '<div class=""upload-result"" style=""background:#f8d7da;color:#842029"">&#50629;&#47196;&#46300; &#49892;&#54056;: ' + e.message + '</div>';
      btn.disabled = false; btn.textContent = '&#50629;&#47196;&#46300;';
    });
}
checkDllStatus();
</script>
</body>
</html>";
            resp.StatusCode = 200;
            resp.ContentType = "text/html; charset=utf-8";
            byte[] buf = Encoding.UTF8.GetBytes(html);
            resp.ContentLength64 = buf.Length;
            resp.OutputStream.Write(buf, 0, buf.Length);
            resp.Close();
        }
    }
}
