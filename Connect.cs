using System;
using System.Reflection;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.CommandBars;

namespace VSBuildMonitor
{
  public struct BuildInfo
  {
    public BuildInfo(string n, long b, long e, bool s)
    {
      name = n;
      begin = b;
      end = e;
      success = s;
    }
    public long begin;
    public long end;
    public string name;
    public bool success;
  }
  /// <summary>The object for implementing an Add-in.</summary>
  /// <seealso class='IDTExtensibility2' />
  public class Connect : IDTExtensibility2, IDTCommandTarget
  {
    public DTE2 _applicationObject;
    public AddIn _addInInstance;
    public DateTime buildTime;
    public OutputWindowPane paneWindow;
    public ChartsControl controlWindow;
    public EnvDTE.BuildEvents buildEvents;
    public EnvDTE.SolutionEvents solutionEvents;
    public Dictionary<string, DateTime> currentBuilds = new Dictionary<string, DateTime>();
    public List<BuildInfo> finishedBuilds = new List<BuildInfo>();
    public static string timeFormat = "HH:mm:ss";
    public static string addinName = "VSBuildMonitor";
    public static string commandToggle = "ToggleCPPH";
    public static string commandFixIncludes = "FixIncludes";
    public static string commandFindReplaceGUIDsInSelection = "FindReplaceGUIDsInSelection";
    public string logFileName;
    public int maxParallelBuilds = 0;
    public int allProjectsCount = 0;
    public int outputCounter = 0;
    public Timer timer = new Timer();

    private void WriteToLog(string line)
    {
      try
      {
        StreamWriter sw = new StreamWriter(logFileName, true);
        sw.Write(line);
        sw.Close();
      }
      catch (Exception)
      {
      }
    }

    /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
    public Connect()
    {
    }

    /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
    /// <param term='application'>Root object of the host application.</param>
    /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
    /// <param term='addInInst'>Object representing this Add-in.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
      _applicationObject = (DTE2)application;
      _addInInstance = (AddIn)addInInst;

      solutionEvents = _applicationObject.Events.SolutionEvents;
      solutionEvents.AfterClosing += new _dispSolutionEvents_AfterClosingEventHandler(solutionEvents_AfterClosing);

      buildEvents = _applicationObject.Events.BuildEvents;

      if (connectMode == ext_ConnectMode.ext_cm_Startup)
      {
        buildEvents.OnBuildBegin += new _dispBuildEvents_OnBuildBeginEventHandler(BuildEvents_OnBuildBegin);
        buildEvents.OnBuildDone += new _dispBuildEvents_OnBuildDoneEventHandler(BuildEvents_OnBuildDone);
        buildEvents.OnBuildProjConfigBegin += new _dispBuildEvents_OnBuildProjConfigBeginEventHandler(BuildEvents_OnBuildProjConfigBegin);
        buildEvents.OnBuildProjConfigDone += new _dispBuildEvents_OnBuildProjConfigDoneEventHandler(BuildEvents_OnBuildProjConfigDone);
      }

      if (connectMode != ext_ConnectMode.ext_cm_CommandLine)
      {
        paneWindow = _applicationObject.ToolWindows.OutputWindow.OutputWindowPanes.Add(addinName);

        string guid = "{bd488241-6ff7-4f10-98b7-e40a1ebbd4ae}";
        Windows2 win = (Windows2)_applicationObject.Windows;
        object ctl = null;
        System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
        Window toolWindow = win.CreateToolWindow2(_addInInstance, asm.Location, addinName + ".ChartsControl", addinName, guid, ref ctl);
        toolWindow.Visible = true;
        //toolWindow.SetTabPicture(Properties.Resources.TabIcon.ToBitmap().GetHbitmap());
        controlWindow = (ChartsControl)ctl;
        controlWindow.host = this;
      }

      if (connectMode == ext_ConnectMode.ext_cm_AfterStartup)
      {
        // The add-in was loaded by hand after startup using the Add-In Manager
        // Initialize it in the same way that when is loaded on startup
        AddMenuCommands();
      }
    }

    void solutionEvents_AfterClosing()
    {
      currentBuilds.Clear();
      finishedBuilds.Clear();
      if (paneWindow != null)
      {
        paneWindow.Clear();
        if (controlWindow != null)
        {
          controlWindow.Refresh();
        }
      }
    }

    void BuildEvents_OnBuildProjConfigDone(string Project, string ProjectConfig, string Platform, string SolutionConfig, bool Success)
    {
      string key = MakeKey(Project, ProjectConfig, Platform);
      if (currentBuilds.ContainsKey(key))
      {
        outputCounter++;
        DateTime start = new DateTime(currentBuilds[key].Ticks - buildTime.Ticks);
        currentBuilds.Remove(key);
        DateTime end = new DateTime(DateTime.Now.Ticks - buildTime.Ticks);
        finishedBuilds.Add(new BuildInfo(key, start.Ticks, end.Ticks, Success));
        TimeSpan s = end - start;
        DateTime t = new DateTime(s.Ticks);
        StringBuilder b = new StringBuilder(outputCounter.ToString("D3"));
        b.Append(" ");
        b.Append(key);
        int space = 50 - key.Length;
        if (space > 0)
        {
          b.Append(' ', space);
        }
        b.Append(" \t");
        b.Append(start.ToString(timeFormat));
        b.Append("\t");
        b.Append(t.ToString(timeFormat));
        b.Append("\n");
        if (paneWindow != null)
        {
          paneWindow.OutputString(b.ToString());
          if (controlWindow != null)
          {
            controlWindow.Refresh();
          }
        }
        else
        {
          WriteToLog(b.ToString());
        }
      }
    }

    string MakeKey(string Project, string ProjectConfig, string Platform)
    {
      FileInfo fi = new FileInfo(Project);
      string key = fi.Name + "|" + ProjectConfig + "|" + Platform;
      return key;
    }

    void BuildEvents_OnBuildProjConfigBegin(string Project, string ProjectConfig, string Platform, string SolutionConfig)
    {
      string key = MakeKey(Project, ProjectConfig, Platform);
      currentBuilds[key] = DateTime.Now;
      if (currentBuilds.Count > maxParallelBuilds)
      {
        maxParallelBuilds = currentBuilds.Count;
      }
    }

    void BuildEvents_OnBuildDone(vsBuildScope Scope, vsBuildAction Action)
    {
      timer.Enabled = false;
      controlWindow.isBuilding = false;
      TimeSpan s = DateTime.Now - buildTime;
      DateTime t = new DateTime(s.Ticks);
      string msg = "Build Total Time: " + t.ToString(timeFormat) + ", max. number of parallel builds: " + maxParallelBuilds.ToString() + "\n";
      if (paneWindow != null)
      {
        _applicationObject.ToolWindows.OutputWindow.ActivePane.OutputString(msg);
        if (controlWindow != null)
        {
          controlWindow.Refresh();
        }
      }
      else
      {
        WriteToLog(msg);
        logFileName = null;
      }
    }

    int GetProjectsCount(Project project)
    {
      int count = 0;
      if (project != null)
      {
        if (project.FullName.ToLower().EndsWith(".vcxproj") ||
            project.FullName.ToLower().EndsWith(".csproj") ||
            project.FullName.ToLower().EndsWith(".vbproj"))
          count = 1;
        if (project.ProjectItems != null)
        {
          foreach (ProjectItem projectItem in project.ProjectItems)
          {
            count += GetProjectsCount(projectItem.SubProject);
          }
        }
      }
      return count;
    }

    void BuildEvents_OnBuildBegin(vsBuildScope Scope, vsBuildAction Action)
    {
      buildTime = DateTime.Now;
      maxParallelBuilds = 0;
      allProjectsCount = 0;
      foreach (Project project in _applicationObject.Solution.Projects)
        allProjectsCount += GetProjectsCount(project);
      outputCounter = 0;
      timer.Enabled = true;
      timer.Interval = 1000;
      timer.Tick += new EventHandler(timer_Tick);
      currentBuilds.Clear();
      finishedBuilds.Clear();
      if (paneWindow != null)
      {
        paneWindow.Clear();
        if (controlWindow != null)
        {
          controlWindow.scrollLast = true;
          controlWindow.isBuilding = true;
          controlWindow.Refresh();
        }
      }
      else
      {
        logFileName = _applicationObject.Solution.FullName;
        if (!string.IsNullOrEmpty(logFileName))
        {
          logFileName += ".timing.csv";
          try
          {
            StreamWriter sw = new StreamWriter(logFileName, false);
            sw.WriteLine("File generated by " + addinName + " Add-In");
            sw.Close();
          }
          catch (Exception)
          {
          }
        }
      }
    }

    public long PercentageProcessorUse()
    {
      long percentage = 0;
      if (maxParallelBuilds > 0)
      {
        long nowTicks = DateTime.Now.Ticks;
        long maxTick = 0;
        long totTicks = 0;
        foreach (BuildInfo info in finishedBuilds)
        {
          totTicks += info.end - info.begin;
          if (info.end > maxTick)
          {
            maxTick = info.end;
          }
        }
        foreach (DateTime start in currentBuilds.Values)
        {
          maxTick = nowTicks - buildTime.Ticks;
          totTicks += nowTicks - start.Ticks;
        }
        totTicks /= maxParallelBuilds;
        if (maxTick > 0)
        {
          percentage = totTicks * 100 / maxTick;
        }
      }
      return percentage;
    }

    void timer_Tick(object sender, EventArgs e)
    {
      if (paneWindow != null)
      {
        if (controlWindow != null)
        {
          controlWindow.Refresh();
        }
      }
    }

    /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
    /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
    {
    }

    /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />		
    public void OnAddInsUpdate(ref Array custom)
    {
    }

    /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnStartupComplete(ref Array custom)
    {
      AddMenuCommands();
    }

    /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
    /// <param term='custom'>Array of parameters that are host application specific.</param>
    /// <seealso class='IDTExtensibility2' />
    public void OnBeginShutdown(ref Array custom)
    {
    }

    /// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
    /// <param term='commandName'>The name of the command to determine state for.</param>
    /// <param term='neededText'>Text that is needed for the command.</param>
    /// <param term='status'>The state of the command in the user interface.</param>
    /// <param term='commandText'>Text requested by the neededText parameter.</param>
    /// <seealso class='Exec' />
    public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
    {
      if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
      {
        if (commandName == addinName + ".Connect." + commandToggle)
        {
          status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
          return;
        }
        if (commandName == addinName + ".Connect." + commandFixIncludes)
        {
          status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
          return;
        }
        if (commandName == addinName + ".Connect." + commandFindReplaceGUIDsInSelection)
        {
          status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
          return;
        }
      }
    }

    void ReplaceIncludes(ProjectItem project)
    {
      if (project.Name.ToLower().EndsWith(".cpp") ||
          project.Name.ToLower().EndsWith(".h") ||
          project.Name.ToLower().EndsWith(".inl") ||
          project.Name.ToLower().EndsWith(".c") ||
          project.Name.ToLower().EndsWith(".hpp"))
      {
        if (project.ContainingProject == null)
          return;
        if (!File.Exists(project.ContainingProject.FullName))
          return;
        FileInfo projectFile = new FileInfo(project.ContainingProject.FullName);
        string fileName = projectFile.DirectoryName + "\\" + project.Name;
        if (File.Exists(fileName))
        {
          FileInfo srcFile = new FileInfo(fileName);
          if (srcFile.IsReadOnly)
            return;
          Dictionary<string, string> replace = new Dictionary<string, string>();
          List<string> lines;
          if (project.Document != null)
          {
            TextSelection ts = project.Document.Selection as TextSelection;
            ts.SelectAll();
            StringBuilder selectedText = new StringBuilder();
            selectedText.Append(ts.Text);
            lines = new List<string>(ts.Text.Split(new char[] { '\r', '\n' }));
          }
          else
          {
            lines = new List<string>(File.ReadAllLines(fileName, Encoding.GetEncoding(1252)));
          }
          foreach (string line in lines)
          {
            string inc = "#include";
            if (line.StartsWith(inc))
            {
              string inlcudedFile = line.Substring(inc.Length).TrimStart(new char[] { ' ', '\t' });
              if (inlcudedFile.Length == 0)
                continue;
              int lBracket = -1;
              int rBracket = -1;
              if (inlcudedFile[0] == '<')
              {
                lBracket = line.IndexOf('<');
                rBracket = line.IndexOf('>', lBracket + 1);
              }
              else if (inlcudedFile[0] == '\"')
              {
                lBracket = line.IndexOf('"');
                rBracket = line.IndexOf('"', lBracket + 1);
              }
              else
              {
                continue;
              }
              if (rBracket <= lBracket)
                continue;
              inlcudedFile = line.Substring(lBracket + 1, rBracket - lBracket - 1);
              if (inlcudedFile.Length <= 0)
                continue;
              if (inlcudedFile[0] == '.')
                continue;
              inlcudedFile = inlcudedFile.Replace('/', '\\');
              string[] directories = inlcudedFile.Split(new char[] { '\\', '/' });
              if (directories.Length == 1)
              {
              }
              else if (directories.Length > 1)
              {
                if (File.Exists(@"f:\c\" + inlcudedFile))
                {
                  string dirUp = "";
                  int subdirs = srcFile.DirectoryName.Split(new char[] { '\\', '/' }).Length - 2;
                  for (int i = 0; i < subdirs; i++)
                  {
                    dirUp += "..\\";
                  }
                  if (File.Exists(srcFile.DirectoryName + "\\" + dirUp + inlcudedFile))
                  {
                    string newLine = line.Substring(0, lBracket) + "\"" + dirUp + inlcudedFile + "\"" + line.Substring(rBracket + 1);
                    if (!replace.ContainsKey(line) && line != newLine)
                      replace.Add(line, newLine);
                  }
                }
              }
            }
          }
          if (replace.Count > 0)
          {
            Document doc = project.Document;
            bool close = false;
            if (doc == null)
            {
              close = true;
              doc = _applicationObject.ItemOperations.OpenFile(fileName).Document;
            }
            TextSelection ts = doc.Selection as TextSelection;
            foreach (var r in replace)
            {
              ts.SelectAll();
              ts.ReplacePattern(r.Key, r.Value);
            }
            if (close)
            {
              doc.Close(vsSaveChanges.vsSaveChangesYes);
            }
            else
            {
              ts.StartOfDocument();
            }
          }
        }
      }
      else
      {
        if (project.ProjectItems != null)
        {
          for (int i = 1; i <= project.ProjectItems.Count; i++)
          {
            ProjectItem projectItem = project.ProjectItems.Item(i);
            ReplaceIncludes(projectItem);
          }
        }
      }
    }

    /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
    /// <param term='commandName'>The name of the command to execute.</param>
    /// <param term='executeOption'>Describes how the command should be run.</param>
    /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
    /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
    /// <param term='handled'>Informs the caller if the command was handled or not.</param>
    /// <seealso class='Exec' />
    public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
    {
      handled = false;
      if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
      {
        if (commandName == addinName + ".Connect." + commandFindReplaceGUIDsInSelection)
        {
          handled = true;

          try
          {
            TextSelection ts = _applicationObject.ActiveDocument.Selection as TextSelection;
            StringBuilder selectedText = new StringBuilder();
            selectedText.Append(ts.Text);
            MatchCollection allGuids = Regex.Matches(ts.Text, "[0-9a-hA-H]+-[0-9a-hA-H]+-[0-9a-hA-H]+-[0-9a-hA-H]+-[0-9a-hA-H]+", RegexOptions.Multiline);
            foreach (Match guid in allGuids)
            {
              selectedText.Replace(guid.Value, System.Guid.NewGuid().ToString().ToUpper());
            }
            ts.Text = selectedText.ToString();
            MessageBox.Show(allGuids.Count.ToString() + " GUIDs replaced.");
          }
          catch (Exception)
          {
          }

          return;
        }
        if (commandName == addinName + ".Connect." + commandFixIncludes)
        {
          handled = true;
          try
          {
            for (int i = 1; i <= _applicationObject.Solution.Projects.Count; i++)
            {
              Project project = _applicationObject.Solution.Projects.Item(i);
              for (int j = 1; j <= project.ProjectItems.Count; j++)
              {
                ProjectItem projectItem = project.ProjectItems.Item(j);
                ReplaceIncludes(projectItem);
                if (projectItem.SubProject != null && projectItem.SubProject.ProjectItems != null)
                {
                  for (int k = 1; k <= projectItem.SubProject.ProjectItems.Count; k++)
                  {
                    ProjectItem subProjectItem = projectItem.SubProject.ProjectItems.Item(k);
                    ReplaceIncludes(subProjectItem);
                  }
                }
              }
            }
          }
          catch (Exception e)
          {
            System.Windows.Forms.MessageBox.Show(e.Message);
          }
        }
        if (commandName == addinName + ".Connect." + commandToggle)
        {
          handled = true;

          try
          {
            string fullName = _applicationObject.ActiveDocument.FullName.ToLower();
            FileInfo fi = new FileInfo(fullName);
            if (fi.Extension == ".cpp" || fi.Extension == ".c")
            {
              string newName = fullName.Replace(fi.Extension, ".h");
              if (File.Exists(newName))
              {
                _applicationObject.ItemOperations.OpenFile(newName);
              }
              else
              {
                newName = fullName.Replace(fi.Extension, ".hpp");
                if (File.Exists(newName))
                  _applicationObject.ItemOperations.OpenFile(newName);
              }
            }
            else if (fi.Extension == ".hpp" || fi.Extension == ".h")
            {
              string newName = fullName.Replace(fi.Extension, ".cpp");
              if (File.Exists(newName))
              {
                _applicationObject.ItemOperations.OpenFile(newName);
              }
              else
              {
                newName = fullName.Replace(fi.Extension, ".c");
                if (File.Exists(newName))
                  _applicationObject.ItemOperations.OpenFile(newName);
              }
            }
          }
          catch (Exception)
          {
          }

          return;
        }
      }
    }

    public void AddMenuCommands()
    {
      try
      {
        CommandBars commandBars = _applicationObject.CommandBars as CommandBars;

        CommandBar menuCommandBar = commandBars["MenuBar"];
        CommandBar toolsCommandBar = GetCommandBarPopup(menuCommandBar, "Tools");

        CommandBarPopup pop = toolsCommandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, toolsCommandBar.Controls.Count + 1, true) as CommandBarPopup;

        pop.CommandBar.Name = addinName;
        pop.Caption = "Robot Dev (" + addinName + ")";
        pop.BeginGroup = true;

        string[] commands = new string[] { commandFindReplaceGUIDsInSelection, commandToggle, commandFixIncludes };
        string[] menuName = new string[] { "Find and Replace GUIDs in selection", "Switch between .cpp and .h", "Fix paths in #include" };

        int pos = 1;

        for (int i = 0; i < commands.Length; i++)
        {
          Command cmd = null;

          try
          {
            cmd = _applicationObject.Commands.Item(_addInInstance.ProgID + "." + commands[i]);
          }
          catch (Exception)
          {
          }

          if (cmd == null)
          {
            cmd = _applicationObject.Commands.AddNamedCommand(_addInInstance, commands[i], commands[i], "", false, 0, null, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled);
          }

          CommandBarButton myCommandBarPopup1Button = cmd.AddControl(pop.CommandBar, pos) as CommandBarButton;

          pos = pos + 1;

          myCommandBarPopup1Button.Caption = menuName[i];
          myCommandBarPopup1Button.BeginGroup = false;
        }

        pop.Visible = true;
      }
      catch (Exception e)
      {
        System.Windows.Forms.MessageBox.Show(addinName + ": " + e.Message);
      }
    }

    CommandBar GetCommandBarPopup(CommandBar parentCommandBar, string commandBarPopupName)
    {
      foreach (CommandBarControl commandBarControl in parentCommandBar.Controls)
      {

        if (commandBarControl.Type == MsoControlType.msoControlPopup)
        {

          CommandBarPopup commandBarPopup = commandBarControl as CommandBarPopup;

          if (commandBarPopup.CommandBar.Name == commandBarPopupName)
          {

            return commandBarPopup.CommandBar;

          }

        }

      }

      return null;
    }

  }
}