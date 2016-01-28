using EmbracePortalDbMgr;
using MtParser.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace MtParser
{
    public partial class Form1 : Form
    {
        private readonly Dictionary<string, Dictionary<string, string>> _directMT2MTDictionary = new Dictionary<string, Dictionary<string, string>>();
        private readonly List<string> _relatedCommonCLS = new List<string>();


        // Dictionary to contain called MT classes and their method in code behind
        // UI class name, MT class name, MT public method name
        // We only store MT public method name, that are called from UT class name.
        private readonly Dictionary<string, Dictionary<string, List<ObjectInfo>>> _uiClassesDictionary = new Dictionary<string, Dictionary<string, List<ObjectInfo>>>();

        // Dictionary for MT class with full path
        // Full path of MT class, MT class content
        private readonly Dictionary<string, string> _cachedMtClassContent = new Dictionary<string, string>();

        // Dictionary to store all stored procedures name for each MT class and method
        // MT class name, MT public method name, stored procedure name.
        // Store all MT public method name in this dictionary
        private readonly Dictionary<string, Dictionary<string, List<ObjectInfo>>> _mtClassesDictionary = new Dictionary<string, Dictionary<string, List<ObjectInfo>>>();

        private const string FacadePattern = ".Facade.";

        public Form1()
        {
            InitClsNotMT();

            InitMT2MTDic();

            InitializeComponent();

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            treeView1.Enabled = false;
            DgvStoredProcedures.Enabled = false;

            TxtSourceFolder.Text = Settings.Default["InitialDirectory"].ToString();
            ImageList iconList = new ImageList();
            iconList.Images.Add(Resources.prj);
            iconList.Images.Add(Resources.folder);
            iconList.Images.Add(Resources.vb);
            iconList.Images.Add(Resources.ukn);
            iconList.Images.Add(Resources.tick);

            treeView1.AfterSelect += treeView1_AfterSelect;
            treeView1.ImageList = iconList;
        }

        private void InitClsNotMT()
        {
            _relatedCommonCLS.Add(".Mail.Email");
            _relatedCommonCLS.Add(".Mail.EmailUtilities");
            _relatedCommonCLS.Add(".Publishers.DBPublisher");
            _relatedCommonCLS.Add(".PDFProcessor.clsPDFProcessor");
            //_relatedCommonCLS.Add("clsSystemParameters");
            _relatedCommonCLS.Add(".Reports.EngageReportServerConnection");
            _relatedCommonCLS.Add(".SiteMaps.PatientPortalSiteMapProvider");
            _relatedCommonCLS.Add("Membership.EmbraceMembershipProvider");
            _relatedCommonCLS.Add(".Membership.PatientPortalMembershipProvider");
            _relatedCommonCLS.Add(".Membership.ProviderPortalMembershipProvider");
            _relatedCommonCLS.Add(".Roles.PatientPortalRoleProvider");

            _mtClassesDictionary[".Mail.Email"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Mail.Email"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Mail.Email"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .Mail.Email CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".Mail.EmailUtilities"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Mail.EmailUtilities"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Mail.EmailUtilities"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .Mail.EmailUtilities CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".Publishers.DBPublisher"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Publishers.DBPublisher"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Publishers.DBPublisher"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN DBPublisher CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".PDFProcessor.clsPDFProcessor"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".PDFProcessor.clsPDFProcessor"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".PDFProcessor.clsPDFProcessor"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .PDFProcessor.clsPDFProcessor CLS, NEED CHECK MANUAL"
            });

            //_mtClassesDictionary["clsSystemParameters"] = new Dictionary<string, List<ObjectInfo>>();
            //_mtClassesDictionary["clsSystemParameters"]["WARNING"] = new List<ObjectInfo>();
            //_mtClassesDictionary["clsSystemParameters"]["WARNING"].Add(new ObjectInfo()
            //{
            //    LineNumber = "---",
            //    Name = "CONTAIN clsSystemParameters CLS, NEED CHECK MANUAL"
            //});

            _mtClassesDictionary[".Reports.EngageReportServerConnection"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Reports.EngageReportServerConnection"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Reports.EngageReportServerConnection"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .Reports.EngageReportServerConnection CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".SiteMaps.PatientPortalSiteMapProvider"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".SiteMaps.PatientPortalSiteMapProvider"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".SiteMaps.PatientPortalSiteMapProvider"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .SiteMaps.PatientPortalSiteMapProvider CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary["Membership.EmbraceMembershipProvider"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary["Membership.EmbraceMembershipProvider"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary["Membership.EmbraceMembershipProvider"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN Membership.EmbraceMembershipProvider CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".Membership.PatientPortalMembershipProvider"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Membership.PatientPortalMembershipProvider"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Membership.PatientPortalMembershipProvider"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .Membership.PatientPortalMembershipProvider CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".Membership.ProviderPortalMembershipProvider"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Membership.ProviderPortalMembershipProvider"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Membership.ProviderPortalMembershipProvider"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .Membership.ProviderPortalMembershipProvider CLS, NEED CHECK MANUAL"
            });

            _mtClassesDictionary[".Roles.PatientPortalRoleProvider"] = new Dictionary<string, List<ObjectInfo>>();
            _mtClassesDictionary[".Roles.PatientPortalRoleProvider"]["WARNING"] = new List<ObjectInfo>();
            _mtClassesDictionary[".Roles.PatientPortalRoleProvider"]["WARNING"].Add(new ObjectInfo()
            {
                LineNumber = "---",
                Name = "CONTAIN .Roles.PatientPortalRoleProvider CLS, NEED CHECK MANUAL"
            });
        }

        private void InitMT2MTDic()
        {
            var clsRuleEngine = new Dictionary<string, string>();
            clsRuleEngine["TestCompileExpression"] = "clsRuleEvaluator";
            clsRuleEngine["CheckRulesForEffects"] = "clsRuleEvaluator";
            clsRuleEngine["GetMemberTokenDataForEvaluator"] = "clsRuleEvaluator";
            clsRuleEngine["SendParticipationStatus"] = "clsMemberDataExchange";
            clsRuleEngine["UpdateHierarchyPaths"] = "clsClientHierarchy";
            clsRuleEngine["GetMemberTokenDataForEvaluator"] = "clsMemberLookup & clsMemberAssessments";
            clsRuleEngine["ProcessSMSEffect"] = "clsDemographics, clsAssessmentMaintenance & clsPatientEncounters";
            clsRuleEngine["ReceiveSMSMessage"] = "clsSystemParameters";
            clsRuleEngine["ReceiveSMSReplyMessage"] = "clsSystemParameters";
            clsRuleEngine["CalculateHPRScorePerMemberDA"] = "clsSystemParameters";
            clsRuleEngine["RunStatEngine"] = "clsRuleEvaluator & clsRuleEngine";
            _directMT2MTDictionary["clsRuleEngine"] = clsRuleEngine;

            var clsClientHierarchy = new Dictionary<string, string>();
            clsClientHierarchy["UpdateClientHierarchy"] = "clsContract";
            _directMT2MTDictionary["clsClientHierarchy"] = clsClientHierarchy;

            var clsLogin = new Dictionary<string, string>();
            clsLogin["Authenticate"] = "clsSystemParameters & clsCrypto";
            _directMT2MTDictionary["clsLogin"] = clsLogin;

            var clsUser = new Dictionary<string, string>();
            clsUser["AuthenticateUser"] = "clsCrypto";
            _directMT2MTDictionary["clsUser"] = clsUser;

            var clsAction = new Dictionary<string, string>();
            clsAction["Create"] = "clsSystemParameters";
            _directMT2MTDictionary["clsAction"] = clsAction;

            var clsStatEngine = new Dictionary<string, string>();
            clsStatEngine["RunStatEngine"] = "clsRuleEvaluator & clsRuleEngine";
            _directMT2MTDictionary["clsStatEngine"] = clsStatEngine;

            var clsHealthCareTeam = new Dictionary<string, string>();
            clsHealthCareTeam["UpdateProvider"] = "clsProviderPortal";
            _directMT2MTDictionary["clsHealthCareTeam"] = clsHealthCareTeam;
        }

        void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            var fileName = Path.GetFileName(treeView1.SelectedNode.Tag + string.Empty);
            if (!_uiClassesDictionary.ContainsKey(fileName))
            {
                DgvStoredProcedures.DataSource = null;
                return;
            }

            List<DataObjectBinding> dObj = new List<DataObjectBinding>();
            foreach (var className in _uiClassesDictionary[fileName].Keys)
            {
                var currentClass = className;

                foreach (var methodName in _uiClassesDictionary[fileName][className])
                {

                    var currentMethod = methodName;

                    if (!_mtClassesDictionary.ContainsKey(className))
                    {
                        MessageBox.Show(className);
                    }
                    else
                    {
                        if (!_mtClassesDictionary[className].ContainsKey(methodName.Name) || _mtClassesDictionary[className][methodName.Name].Count == 0)
                        {
                            dObj.Add(new DataObjectBinding(fileName, currentClass, currentMethod.Name, currentMethod.LineNumber, string.Empty, string.Empty));
                        }
                        else
                        {
                            foreach (var storedProcedureName in _mtClassesDictionary[className][methodName.Name])
                            {
                                bool hasAny = false;
                                if (_filterStoredProcedure.Count == 0 || _filterStoredProcedure.Contains(storedProcedureName.Name))
                                {
                                    hasAny = true;
                                    dObj.Add(new DataObjectBinding(fileName, currentClass, currentMethod.Name, currentMethod.LineNumber, storedProcedureName.Name, storedProcedureName.LineNumber));
                                }

                                if (!hasAny)
                                {
                                    dObj.Add(new DataObjectBinding(fileName, currentClass, currentMethod.Name, currentMethod.LineNumber, string.Empty, string.Empty));
                                }
                            }
                        }
                    }
                }
            }

            DgvStoredProcedures.DataSource = dObj;
        }

        string _solutionFolder;

        private void BtnLookupDependencies_Click(object sender, EventArgs e)
        {
            var eLeng = "PatientPortal\\EmbracePortalUI\\EmbracePortalUI.vbproj".Length;
            var temp = TxtSourceFolder.Text.Length - eLeng;
            if (temp <= 0)
            {
                MessageBox.Show("Invalid source path!");
                return;
            }

            _solutionFolder = TxtSourceFolder.Text.Substring(0, temp);

            toolStripProgressBar1.MarqueeAnimationSpeed = 30;
            toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            toolStripLabel1.Text = "Parsing...";

            Task t = Task.Factory.StartNew(() =>
            {
                IterateUiLayerFolder();

                IterateMtClassesToGetSpNames();
            });

            // ReSharper disable ImplicitlyCapturedClosure
            t.ContinueWith(success =>
                // ReSharper restore ImplicitlyCapturedClosure
                // Export to excel
                Invoke((MethodInvoker)delegate
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    //button3.Enabled = true;
                    button4.Enabled = true;
                    treeView1.Enabled = true;
                    DgvStoredProcedures.Enabled = true;
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Parse successfully.";
                }), TaskContinuationOptions.NotOnFaulted);
            t.ContinueWith(fail =>
            {
                //log the exception i.e.: Fail.Exception.InnerException);
                Invoke((MethodInvoker)delegate
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    //button3.Enabled = true;
                    button4.Enabled = true;
                    treeView1.Enabled = true;
                    DgvStoredProcedures.Enabled = true;
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Error";
                });
            }, TaskContinuationOptions.OnlyOnFaulted);
        }

        private void IterateMtClassesToGetSpNames()
        {
            foreach (var fileName in _uiClassesDictionary.Keys)
            {
                foreach (var className in _uiClassesDictionary[fileName].Keys)
                {
                    Invoke((MethodInvoker)delegate
                     {
                         toolStripLabel1.Text = "Parsing " + fileName + "\\" + className;
                     });

                    var mtClassContent = GetMtClassContent(className);

                    if (string.IsNullOrEmpty(mtClassContent))
                    {
                        continue;
                    }

                    if (!_mtClassesDictionary.ContainsKey(className))
                    {
                        _mtClassesDictionary.Add(className, new Dictionary<string, List<ObjectInfo>>());

                        // Parse all method in current MT class first.
                        GetStoredProcedureNameInAllMtMethods(mtClassContent, className);

                        // Add all SPs from privated method that was called in public method.
                        MergeSpsNameListFromPrivateMethods(mtClassContent, className);
                    }
                }
            }
        }

        private void GetStoredProcedureNameInAllMtMethods(string classContent, string className)
        {
            var lines = classContent.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();

            var i = 0;
            var isCheckingStoredProcedure = false;
            string methodName = string.Empty;

            while (i < lines.Count)
            {
                var pureLine = lines[i++].Trim();
                if (pureLine.StartsWith("'")) continue;

                // Check reference to Other class different to MT
                foreach (var cl in _relatedCommonCLS)
                {
                    if (pureLine.Contains(cl) || pureLine.Contains("New " + cl.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries)[1]))
                    {
                        _mtClassesDictionary[className][methodName].Add(_mtClassesDictionary[cl]["WARNING"][0]);
                    }
                }

                if (pureLine.StartsWith("Public Function") ||
                    pureLine.StartsWith("Public Sub") ||
                    pureLine.StartsWith("Public Shared Function") ||
                    pureLine.StartsWith("Public Shared Sub") ||
                    pureLine.StartsWith("Private Shared Function") ||
                    pureLine.StartsWith("Private Shared Sub") ||
                    pureLine.StartsWith("Private Function") ||
                    pureLine.StartsWith("Private Sub") ||
                    pureLine.StartsWith("Private Function") ||
                    pureLine.StartsWith("Private Sub") ||
                    pureLine.StartsWith("Sub") ||
                    pureLine.StartsWith("Function"))
                {
                    // Turn on flag to start look for stored procedure within current method.
                    isCheckingStoredProcedure = true;

                    // Get method name
                    methodName = pureLine.Substring(0, pureLine.IndexOf("(", StringComparison.Ordinal));
                    methodName = methodName.Substring(methodName.LastIndexOf(" ", StringComparison.Ordinal) + 1);

                    if (!_mtClassesDictionary[className].ContainsKey(methodName))
                    {
                        _mtClassesDictionary[className].Add(methodName, new List<ObjectInfo>());
                        if (_directMT2MTDictionary.ContainsKey(className) && _directMT2MTDictionary[className].ContainsKey(methodName))
                        {
                            var warningMsg = "WARNING! CHECK IN METHOD: " + className + "." + methodName + "() FOR CLASS: " + _directMT2MTDictionary[className][methodName];

                            var sp = new ObjectInfo();
                            sp.Name = warningMsg;
                            sp.LineNumber = "---";

                            _mtClassesDictionary[className][methodName].Add(sp);
                        }
                    }
                }

                // Check only within scope
                if (isCheckingStoredProcedure)
                {
                    if (pureLine.Contains(".CommandText"))
                    {
                        string storedProcedureName;

                        if (pureLine.IndexOf("\"", StringComparison.Ordinal) < 0)
                        {
                            storedProcedureName = pureLine.Substring(pureLine.IndexOf("=", StringComparison.Ordinal) + 1).Trim();
                        }
                        else
                        {
                            storedProcedureName = pureLine.Substring(pureLine.IndexOf("\"", StringComparison.Ordinal) + 1).Trim();
                            storedProcedureName = storedProcedureName.Substring(0, storedProcedureName.Length - 1);
                        }
                        ObjectInfo sp = new ObjectInfo();
                        sp.Name = storedProcedureName;
                        sp.LineNumber = i.ToString();
                        _mtClassesDictionary[className][methodName].Add(sp);
                    }
                    else if (pureLine.Contains("SqlCommand(") && !pureLine.Contains("SqlCommand()"))
                    {
                        var parameter = pureLine.Substring(pureLine.IndexOf("(") + 1, pureLine.LastIndexOf(")") - 1 - pureLine.IndexOf("(") + 1);

                        if (parameter.Contains(","))
                        {
                            parameter = parameter.Split(',')[0];
                        }

                        var oi = new ObjectInfo();
                        oi.Name = parameter;
                        oi.LineNumber = i.ToString();

                        _mtClassesDictionary[className][methodName].Add(oi);
                    }
                }

                if (isCheckingStoredProcedure && (pureLine.StartsWith("End Function") || pureLine.StartsWith("End Sub")))
                {
                    // End of scope of current method.
                    isCheckingStoredProcedure = false;
                }
            }
        }

        private void MergeSpsNameListFromPrivateMethods(string classContent, string className)
        {
            var lines = classContent.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();

            var i = 0;
            string publicMethodName = string.Empty;

            while (i < lines.Count)
            {
                var pureLine = lines[i++].Trim();
                if (pureLine.StartsWith("'")) continue;

                if (pureLine.StartsWith("Public Function") ||
                    pureLine.StartsWith("Public Sub") ||
                    pureLine.StartsWith("Public Shared Function") ||
                    pureLine.StartsWith("Public Shared Sub") ||
                    pureLine.StartsWith("Private Shared Function") ||
                    pureLine.StartsWith("Private Shared Sub") ||
                    pureLine.StartsWith("Private Function") ||
                    pureLine.StartsWith("Private Sub") ||
                    pureLine.StartsWith("Private Function") ||
                    pureLine.StartsWith("Private Sub") ||
                    pureLine.StartsWith("Sub") ||
                    pureLine.StartsWith("Function"))
                {

                    // Get method name
                    publicMethodName = pureLine.Substring(0, pureLine.IndexOf("(", StringComparison.Ordinal));
                    publicMethodName = publicMethodName.Substring(publicMethodName.LastIndexOf(" ", StringComparison.Ordinal) + 1);

                    var methodLines = GetMethodContent(lines, publicMethodName);

                    List<string> checkedMethods = new List<string>();
                    checkedMethods.Add(publicMethodName);
                    GetSPFromPrivateMethod(className, publicMethodName, publicMethodName, methodLines, lines, ref checkedMethods);
                }
            }
        }

        private void GetSPFromPrivateMethod(string className, string rootMethod, string currentMethodName, List<string> methodLines, List<string> clsContent, ref List<string> addedMethods)
        {
            foreach (var ml in methodLines)
            {
                // Parse all private method to merge into current public method.
                // But currently we don't distinguish private and public here.
                foreach (var privateMethodName in _mtClassesDictionary[className].Keys)
                {
                    if (addedMethods.Contains(privateMethodName))
                    {
                        continue;
                    }

                    if (ml.Contains(string.Format("{0}(", privateMethodName)))
                    {
                        addedMethods.Add(privateMethodName);
                        // Add SP name from private class into current public method.
                        foreach (var sp in _mtClassesDictionary[className][privateMethodName])
                        {
                            if (!_mtClassesDictionary[className][rootMethod].Contains(sp))
                                _mtClassesDictionary[className][rootMethod].Add(sp);
                        }

                        var privateMethodLines = GetMethodContent(clsContent, privateMethodName);

                        GetSPFromPrivateMethod(className, rootMethod, privateMethodName, privateMethodLines, clsContent, ref addedMethods);
                    }
                }
            }
        }

        private List<string> GetMethodContent(List<string> lines, string publicMethodName)
        {
            List<string> content = new List<string>();
            var isCheckingStoredProcedure = false;
            string tempPubname = null;
            var ignoreFirstLine = false;

            foreach (var line in lines)
            {
                var pureLine = line.Trim();

                if (pureLine.StartsWith("'")) continue;

                if (pureLine.StartsWith("Public Function") ||
                    pureLine.StartsWith("Public Sub") ||
                    pureLine.StartsWith("Public Shared Function") ||
                    pureLine.StartsWith("Public Shared Sub") ||
                    pureLine.StartsWith("Private Shared Function") ||
                    pureLine.StartsWith("Private Shared Sub") ||
                    pureLine.StartsWith("Private Function") ||
                    pureLine.StartsWith("Private Sub") ||
                    pureLine.StartsWith("Private Function") ||
                    pureLine.StartsWith("Private Sub") ||
                    pureLine.StartsWith("Sub") ||
                    pureLine.StartsWith("Function"))
                {
                    // Get method name
                    tempPubname = pureLine.Substring(0, pureLine.IndexOf("(", StringComparison.Ordinal));
                    tempPubname = tempPubname.Substring(tempPubname.LastIndexOf(" ", StringComparison.Ordinal) + 1);

                    if (tempPubname == publicMethodName)
                    {
                        // Turn on flag to start look for stored procedure within current method.
                        isCheckingStoredProcedure = true;
                    }

                    ignoreFirstLine = true;
                }

                if (isCheckingStoredProcedure && (pureLine.StartsWith("End Function") || pureLine.StartsWith("End Sub")))
                {
                    // End of scope of current method.
                    break;
                }

                if (!ignoreFirstLine && isCheckingStoredProcedure)
                {
                    content.Add(pureLine);
                }

                ignoreFirstLine = false;
            }

            return content;
        }

        private string GetMtClassContent(string mtClassName)
        {
            // Check in cache first.
            if (_cachedMtClassContent.ContainsKey(mtClassName))
            {
                using (var sr = new StreamReader(_cachedMtClassContent[mtClassName]))
                {
                    return sr.ReadToEnd();
                }
            }

            foreach (var filePath in Directory.EnumerateFiles(_solutionFolder, "*.vb", SearchOption.AllDirectories))
            {
                if (filePath.EndsWith(string.Format(@"MT\{0}.vb", mtClassName)))
                {
                    using (var sr = new StreamReader(filePath))
                    {
                        // Add for next usages.
                        _cachedMtClassContent[mtClassName] = filePath;
                        return sr.ReadToEnd();
                    }
                }
            }

            return string.Empty;
        }

        TreeNodeSimulation _projNode;
        private void IterateUiLayerFolder()
        {
            var basedPath = Path.GetDirectoryName(TxtSourceFolder.Text);

            var doc = new XmlDocument();
            try
            {
                doc.Load(TxtSourceFolder.Text);
            }
            catch
            {
                MessageBox.Show("Cant load project file");
                return;
            }


            _projNode = new TreeNodeSimulation(Path.GetFileName(TxtSourceFolder.Text), TxtSourceFolder.Text, 0);

            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("ms", "http://schemas.microsoft.com/developer/msbuild/2003");
            var includedFiles = doc.SelectNodes(@"//ms:Compile[@Include] | //ms:Content[@Include]", nsmgr);
            var total = includedFiles.Count;
            int index = 1;
            foreach (var f in includedFiles)
            {
                Invoke((MethodInvoker)delegate
                {
                    toolStripLabel1.Text = "Loading " + index++ + "/" + total + " files.";
                });

                var fnode = f as XmlNode;
                var n = fnode.Attributes.GetNamedItem("Include");
                var temp = n.Value;

                var segments = temp.Split(new[] { "\\" }, StringSplitOptions.None);
                if (Path.GetExtension(segments[segments.Length - 1]).ToLower() == ".vb" &&
                    (segments[segments.Length - 1].Substring(segments[segments.Length - 1].Length - 12, 12).ToLower() != ".designer.vb" &&
                    segments[segments.Length - 1].ToLower() != "assemblyinfo.vb"))
                {
                    if (segments.Length == 1)
                    {

                        var fNode = new TreeNodeSimulation(segments[0], Path.Combine(basedPath, temp), GetImageIndex(segments[0]));
                        _projNode.Nodes.Add(fNode);

                        ParseUiClassToGetMtClassesAndMtMethodNames(fNode.Tag);
                    }
                    else
                    {
                        var parN = _projNode;

                        for (var i = 0; i < segments.Length - 1; i++)
                        {
                            var curN = parN.Nodes.FirstOrDefault(p => p.Name == segments[i]);

                            if (curN == null)
                            {

                                var currPath = string.Empty;
                                for (var j = 0; j <= i; j++)
                                {
                                    currPath += segments[j] + "\\";
                                }

                                curN = new TreeNodeSimulation(segments[i], Path.Combine(basedPath, currPath), 1);
                                parN.Nodes.Add(curN);
                            }

                            parN = curN;
                        }

                        var fileName = segments[segments.Length - 1];
                        var leafN = new TreeNodeSimulation(fileName, Path.Combine(basedPath, temp), GetImageIndex(fileName));
                        parN.Nodes.Add(leafN);

                        ParseUiClassToGetMtClassesAndMtMethodNames(leafN.Tag);
                    }
                }
            }

            SortNodes(_projNode);

            Invoke((MethodInvoker)delegate
                    {
                        BindTree(_projNode, treeView1.Nodes);

                        if (treeView1.Nodes.Count > 0)
                            treeView1.Nodes[0].Expand();

                        toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                        toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                        toolStripLabel1.Text = "Loaded files";
                    });
        }

        private void SortNodes(TreeNodeSimulation tns)
        {
            if (tns.Nodes.Count == 0)
            {
                return;
            }

            tns.Nodes.Sort((p1, p2) => string.Compare(p1.Name, p2.Name, true));

            foreach (var t in tns.Nodes)
            {
                SortNodes(t);
            }
        }

        private void BindTree(TreeNodeSimulation visual, TreeNodeCollection nodes)
        {
            var n = new TreeNode(visual.Name);
            n.Name = visual.Name;
            n.Tag = visual.Tag;
            n.SelectedImageIndex = 4;
            n.ImageIndex = visual.ImageIndex;

            nodes.Add(n);

            if (visual.Nodes.Count == 0)
                return;

            foreach (var child in visual.Nodes)
            {
                BindTree(child, n.Nodes);
            }
        }

        private int GetImageIndex(string fileName)
        {
            var index = fileName.LastIndexOf(".");
            var extension = fileName.Substring(index, fileName.Length - index);

            switch (extension.ToLower())
            {
                case ".vb":
                    return 2;

                default:
                    return 3;
            }
        }

        private void ParseUiClassToGetMtClassesAndMtMethodNames(string filePath)
        {
            using (var sr = new StreamReader(filePath))
            {
                var fileContent = sr.ReadToEnd();
                var fileName = Path.GetFileName(filePath);

                var lines = fileContent.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
                {
                    var pureLine = lines[lineIndex].Trim();
                    if (pureLine.StartsWith("'")) continue;

                    // Check reference to Other class different to MT
                    foreach (var cl in _relatedCommonCLS)
                    {
                        if (pureLine.Contains(cl) || pureLine.Contains("New " + cl.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries)[1]))
                        {
                            if (!_uiClassesDictionary.ContainsKey(fileName))
                            {
                                _uiClassesDictionary.Add(fileName, new Dictionary<string, List<ObjectInfo>>());
                                // Check add MT class name
                                if (!_uiClassesDictionary[fileName].ContainsKey(cl))
                                {
                                    _uiClassesDictionary[fileName].Add(cl, new List<ObjectInfo>());
                                    _uiClassesDictionary[fileName][cl].Add(new ObjectInfo() { LineNumber = "---", Name = "WARNING" });
                                }
                            }
                        }
                    }

                    // Check contain facade
                    if (pureLine.Contains(FacadePattern))
                    {

                        // Instanciate value for current file (contains facade).
                        if (!_uiClassesDictionary.ContainsKey(fileName))
                        {
                            _uiClassesDictionary.Add(fileName, new Dictionary<string, List<ObjectInfo>>());
                        }

                        // Check last "." to avoid they initalize again like
                        // ...
                        var mtClassName = pureLine.Substring(pureLine.LastIndexOf(".", StringComparison.Ordinal) + 1).Trim();

                        if (mtClassName.LastIndexOf("()", StringComparison.Ordinal) > 0)
                        {
                            mtClassName = mtClassName.Substring(0, mtClassName.Length - 2);
                        }

                        // Check some special cases whose MT class name is not consistent.
                        if ("clsInspirations".Equals(mtClassName))
                        {
                            mtClassName = "clsInspiration";
                        }

                        if ("clsPatientPortalAdmin".Equals(mtClassName))
                        {
                            mtClassName = "clsPatienPortalAdmin";
                        }

                        if ("clsLab".Equals(mtClassName))
                        {
                            mtClassName = "clsLabData";
                        }

                        if ("clsProvider".Equals(mtClassName))
                        {
                            continue;
                        }

                        if ("clsAllergy".Equals(mtClassName))
                        {
                            mtClassName = "clsAllergyMT";
                        }

                        if ("clsNotesConsolidation".Equals(mtClassName))
                        {
                            mtClassName = "clsNotesConsolidationData";
                        }

                        // Check add MT class name
                        if (!_uiClassesDictionary[fileName].ContainsKey(mtClassName))
                        {
                            _uiClassesDictionary[fileName].Add(mtClassName, new List<ObjectInfo>());
                        }

                        List<ObjectInfo> methods = null;

                        // Check class variable
                        if (pureLine.Trim().StartsWith("Private"))
                        {
                            var variableName = pureLine.Substring(8, pureLine.IndexOf(" As ", StringComparison.OrdinalIgnoreCase) - 8).Trim();
                            if (variableName.Contains(" "))
                            {
                                methods = GetCalledMtFunctionsNameInsideUILayer(lines, lineIndex, variableName.Split(' ')[1], true);
                            }
                            else
                            {
                                methods = GetCalledMtFunctionsNameInsideUILayer(lines, lineIndex, variableName, true);
                            }
                        }
                        else if (pureLine.Trim().StartsWith("Dim")) // Check method variable
                        {
                            var variableName = pureLine.Substring(4, pureLine.IndexOf(" As ", StringComparison.OrdinalIgnoreCase) - 4).Trim();

                            methods = GetCalledMtFunctionsNameInsideUILayer(lines, lineIndex, variableName, false);
                        }
                        else if (pureLine.Trim().StartsWith("Using")) // Check method variable
                        {
                            var variableName = pureLine.Substring(5, pureLine.IndexOf(" As ", StringComparison.OrdinalIgnoreCase) - 5).Trim();

                            methods = GetCalledMtFunctionsNameInsideUILayer(lines, lineIndex, variableName, false, true);
                        }

                        // Add result into dictionary.
                        if (methods != null)
                        {
                            foreach (var method in methods)
                            {
                                if (!_uiClassesDictionary[fileName][mtClassName].Any(p=>p.Name == method.Name))
                                {
                                    _uiClassesDictionary[fileName][mtClassName].Add(method);
                                }
                            }
                        }
                    }
                }
            }
        }


        private List<ObjectInfo> GetCalledMtFunctionsNameInsideUILayer(List<string> lines, int lineIndex, string variableName, bool isClassVariable, bool isUsingVariable = false)
        {
            var result = new List<ObjectInfo>();
            bool isInWith = false;

            for (int i = lineIndex + 1; i < lines.Count; i++)
            {
                var pureLine = lines[i].Trim();
                if (pureLine.StartsWith("'"))
                {
                    continue;
                }

                // Check variable name within method scope
                if (!isClassVariable)
                {
                    if (pureLine.StartsWith("End Function") || pureLine.StartsWith("End Sub"))
                    {
                        break;
                    }
                }

                if (isUsingVariable)
                {
                    if (pureLine.StartsWith("End Using"))
                    {
                        break;
                    }
                }

                if (isInWith)
                {
                    if ("End With".Equals(pureLine))
                    {
                        isInWith = false;
                    }
                }

                if (pureLine.StartsWith(string.Format("With {0}", variableName)))
                {
                    isInWith = true;
                }

                if (isInWith)
                {
                    var si = -1;
                    var oi = -1;
                    var devia = 1;
                    if (pureLine.StartsWith("."))
                    {
                        si = 0;
                        oi = pureLine.IndexOf("(", StringComparison.Ordinal);
                    }

                    if (pureLine.Contains(" ."))
                    {
                        si = pureLine.IndexOf(" .", StringComparison.Ordinal);

                        oi = pureLine.IndexOf("(", si, StringComparison.Ordinal);

                        devia = 2;
                    }

                    if (si == -1 || oi <= si)
                    {
                        continue;
                    }

                    if (si == -1) continue;
                    var mn = pureLine.Substring(si + devia,
                                                   oi - (si + devia));

                    if (mn.Contains("=") || mn.Contains(".") || mn.Contains(" ") || mn.Contains(","))
                    {
                        continue;
                    }

                    if (mn != "UseDefaultLocation" && mn != "Dispose")
                    {
                        ObjectInfo method = new ObjectInfo();
                        method.Name = mn;
                        method.LineNumber = i.ToString();

                        result.Add(method);
                    }
                }

                var facadeStartedIndex = pureLine.IndexOf(string.Format("{0}.", variableName), StringComparison.Ordinal);
                if (facadeStartedIndex < 0)
                {
                    continue;
                }

                var openBracketIndex = pureLine.IndexOf("(", facadeStartedIndex, StringComparison.Ordinal);

                if (openBracketIndex < 0)
                {
                    continue;
                }

                var methodName = pureLine.Substring(facadeStartedIndex + variableName.Length + 1,
                                                    openBracketIndex - (facadeStartedIndex + variableName.Length + 1));


                if (methodName.Contains("=") || methodName.Contains(".") || methodName.Contains(" ") || methodName.Contains(","))
                {
                    continue;
                }

                if (methodName != "UseDefaultLocation" && methodName != "Dispose")
                {
                    ObjectInfo method = new ObjectInfo();
                    method.Name = methodName;
                    method.LineNumber = i.ToString();

                    result.Add(method);
                }
            }

            return result;
        }

        private void WriteToExcelFile()
        {
            var storedProcedureNames = new List<string>();

            var xlApp = new Microsoft.Office.Interop.Excel.Application { Visible = true };

            var wb = xlApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);

            // List all store procedure and their LOC.
            var ws = wb.Worksheets.Add();
            ws.Move(After: wb.Sheets[wb.Sheets.Count]);

            ws.Name = "MT class names";

            // Write header for each sheet
            ws.Cells[1, 1] = "File Name";
            ws.Cells[1, 2] = "MT Class name";
            ws.Cells[1, 3] = "Method name";
            ws.Cells[1, 4] = "Method line nubmer";
            ws.Cells[1, 5] = "Stored Procedure name";
            ws.Cells[1, 6] = "Stored Procedure number";
            var i = 2;
            foreach (var fileName in _uiClassesDictionary.Keys)
            {
                ws.Cells[i, 1] = fileName;

                foreach (var className in _uiClassesDictionary[fileName].Keys)
                {
                    ws.Cells[i, 2] = className;

                    foreach (var method in _uiClassesDictionary[fileName][className])
                    {

                        ws.Cells[i, 3] = method.Name;
                        ws.Cells[i, 4] = method.LineNumber;
                        if (!_mtClassesDictionary.ContainsKey(className))
                        {
                            MessageBox.Show(className);
                        }
                        else
                        {
                            if (!_mtClassesDictionary[className].ContainsKey(method.Name) || _mtClassesDictionary[className][method.Name].Count == 0)
                            {
                                ws.Cells[i, 5] = string.Empty;
                                ws.Cells[i++, 6] = string.Empty;
                            }
                            else
                            {
                                foreach (var storedProcedureName in _mtClassesDictionary[className][method.Name])
                                {
                                    bool hasAny = false;

                                    if (_filterStoredProcedure.Count == 0 || _filterStoredProcedure.Contains(storedProcedureName.Name))
                                    {
                                        hasAny = true;
                                        ws.Cells[i, 5] = storedProcedureName.Name;
                                        ws.Cells[i++, 6] = storedProcedureName.LineNumber;
                                        if (!storedProcedureNames.Contains(storedProcedureName.Name))
                                        {
                                            storedProcedureNames.Add(storedProcedureName.Name);
                                        }
                                    }

                                    if (!hasAny)
                                    {
                                        ws.Cells[i, 5] = string.Empty;
                                        ws.Cells[i++, 6] = string.Empty;
                                    }
                                }
                            }
                        }


                    }
                }
            }

            // Auto adjust column width
            ws.UsedRange.Columns.AutoFit();

            storedProcedureNames.Sort();

            // List all store procedure and their LOC.
            ws = wb.Worksheets.Add();
            ws.Move(After: wb.Sheets[wb.Sheets.Count]);

            ws.Name = "StoredProcedureList";

            // Write header for each sheet
            ws.Cells[1, 1] = "Stored Procedure Name";

            i = 2;
            foreach (var storedProcedureName in storedProcedureNames)
            {
                ws.Cells[i++, 1] = storedProcedureName;
            }

            ws.UsedRange.Columns.AutoFit();
        }

        private void LstTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            //DgvStoredProcedures.DataSource = _tableAndSpsDictionary[LstTables.SelectedItem.ToString()];
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Settings.Default["InitialDirectory"].ToString();

            openFileDialog1.Filter = "Project File (*.vbproj)|*.vbproj";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                TxtSourceFolder.Text = openFileDialog1.FileName;
                Settings.Default["InitialDirectory"] = openFileDialog1.FileName;
                Settings.Default.Save();
            }
        }

        private void WriteExcel_Clk(object sender, EventArgs e)
        {
            toolStripProgressBar1.MarqueeAnimationSpeed = 30;
            toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            toolStripLabel1.Text = "Writing...";

            Task t = Task.Factory.StartNew(() =>
            {

                WriteToExcelFile();

                Invoke((MethodInvoker)delegate
                {
                    toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    toolStripProgressBar1.Style = ProgressBarStyle.Blocks;
                    toolStripLabel1.Text = "Write successfully.";
                });
            });

        }

        List<string> _filterStoredProcedure = new List<string>();

        private void FilterLocalizedSP_Clk(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Settings.Default["ExcelPath"].ToString();

            openFileDialog1.Filter = "Excel File (*.xlsx)|*.xlsx";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Settings.Default["ExcelPath"] = Path.GetDirectoryName(openFileDialog1.FileName);
                Settings.Default.Save();

                Microsoft.Office.Interop.Excel.Application appExl;
                Microsoft.Office.Interop.Excel.Workbook workbook;
                Microsoft.Office.Interop.Excel.Worksheet NwSheet;
                Microsoft.Office.Interop.Excel.Range ShtRange;
                appExl = new Microsoft.Office.Interop.Excel.Application();


                //Opening Excel file(myData.xlsx)
                workbook = appExl.Workbooks.Open(openFileDialog1.FileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                NwSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.get_Item(1);

                int Rnum = 0;

                ShtRange = NwSheet.UsedRange; //gives the used cells in sheet

                for (Rnum = 1; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    _filterStoredProcedure.Add((ShtRange.Cells[Rnum, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                }

                workbook.Close(true, Missing.Value, Missing.Value);
                appExl.Quit();
            }
        }

        private void ShowSrc_Clk(object sender, EventArgs e)
        {
            var temp = DgvStoredProcedures.DataSource as List<DataObjectBinding>;

            if (temp == null || temp.Count == 0)
            {
                MessageBox.Show("No class!");
                return;
            }

            var checkDuplication = new List<string>();

            foreach (var item in temp)
            {
                if (checkDuplication.Contains(item.ClassName))
                {
                    continue;
                }

                checkDuplication.Add(item.ClassName);

                foreach (var filePath in Directory.EnumerateFiles(_solutionFolder, "*.vb", SearchOption.AllDirectories))
                {
                    if (filePath.EndsWith(string.Format(@"MT\{0}.vb", item.ClassName)))
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                }
            }
        }

        private void UpdateMT_Clk(object sender, EventArgs e)
        {
            MessageBox.Show("Not Implemented yet!");
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode != null)
                System.Diagnostics.Process.Start(treeView1.SelectedNode.Tag + string.Empty);
        }
    }
}
