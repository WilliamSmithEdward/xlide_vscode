# An Office application's macro Security Notice, accepted for a test harness.
#
# Office can stop a file's load with a modal "Microsoft <Application>
# Security Notice" before the call that opened it returns, and
# AutomationSecurity 1 (msoAutomationSecurityLow) does not always prevent it.
# Measured in Microsoft 365 (16.0 build 20527) from the temp folder:
# PowerPoint raises one for every add-in a project references; Word for a
# template a project references while the VBE is shown; Excel and Word for a
# file opened while the VBE is shown with AutomationSecurity 2; Access for
# none of those, nor for a database marked as downloaded. A harness that does
# not answer it waits for its own timeout, so the harnesses that drive Office
# load this file and accept the notice whenever one appears, without ever
# waiting for one that does not.
#
# Dot-sourced, it defines the type XlideSecurityNotice and nothing else, so it
# changes no variable of the script that loads it:
#   - Accept(pid[, folder]): presses the enabling button of each notice that
#     process (any process, for pid 0) shows now, and returns one line per
#     notice. Given a folder, a notice that names no file in it is left alone.
#   - Poll(pid, folder): the same for notices not already answered, for a
#     caller that asks again and again.
#   - Start(pid) / Settle(ms) / Stop() / Accepted: Poll every 100 ms on a
#     background thread, while the caller waits in a COM call that may raise
#     a notice. Settle waits only while a notice is being answered, so that
#     its line is in Accepted; with none, it returns at once.
#
# The notice is not a top-level UI Automation element, and PowerPoint's has no
# child windows for its buttons, so it is found by window enumeration and its
# buttons by UI Automation from its handle. Only the test harnesses use this:
# XLIDE itself never answers a Security Notice in someone's Office.

if (-not ('XlideSecurityNotice' -as [type])) {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, WindowsBase
    Add-Type -ReferencedAssemblies @(
        [System.Windows.Automation.AutomationElement].Assembly.Location,
        [System.Windows.Automation.ControlType].Assembly.Location,
        [System.Windows.Rect].Assembly.Location
    ) -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

public static class XlideSecurityNotice {
    // The button that lets the file load with its code: Excel, Word and
    // PowerPoint say Enable Macros. Access raised no notice in any way
    // measured; its Security Notice asks to Open or Cancel.
    private static readonly string[] AcceptButtons = { "Enable Macros", "Open" };

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern uint GetLongPathName(string shortPath, StringBuilder longPath, uint size);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern uint GetShortPathName(string longPath, StringBuilder shortPath, uint size);

    // What one caller that asks again and again knows: the notices it has
    // answered (pressed or left alone) and the ones it has reported, while
    // each stays up, and a lock held while it answers and reports them.
    private sealed class Watch {
        public readonly HashSet<IntPtr> Done = new HashSet<IntPtr>();
        public readonly HashSet<IntPtr> Reported = new HashSet<IntPtr>();
        public readonly object Busy = new object();
    }

    private static readonly object gate = new object();
    private static readonly List<string> accepted = new List<string>();
    private static readonly Watch polled = new Watch();
    private static Watch running;
    private static ManualResetEvent stopEvent;
    private static Thread watcherThread;
    private static int generation;

    public static bool IsNoticeTitle(string title) {
        return title != null && title.StartsWith("Microsoft ") && title.EndsWith(" Security Notice");
    }

    public static List<string> Accept(uint processId) {
        return Accept(processId, null);
    }

    // One line per notice: its title, what was done, and what the notice
    // says, its file path included. A notice with no known button, or whose
    // button could not be pressed, is left for the harness's timeout to
    // report.
    public static List<string> Accept(uint processId, string folder) {
        List<string> lines = new List<string>();
        foreach (IntPtr hWnd in FindNotices(processId)) {
            bool settled;
            lines.Add(Answer(hWnd, folder, out settled));
        }
        return lines;
    }

    // As Accept, for a caller that asks again and again: a notice is
    // answered and reported once while it stays up, and a new one, even with
    // the same words, is a window of its own. One not pressed, whether left
    // alone or not readable or pressable yet, as while Office is still
    // building it, is reported once and looked at again on every later call.
    public static List<string> Poll(uint processId, string folder) {
        lock (polled) {
            return PollInto(polled, processId, folder, null);
        }
    }

    private static List<string> PollInto(Watch watch, uint processId, string folder, Func<bool> current) {
        List<IntPtr> found = FindNotices(processId);
        // A handle no longer showing a notice is forgotten, so a new notice
        // that is given it again is answered.
        watch.Done.IntersectWith(found);
        watch.Reported.IntersectWith(found);
        List<string> lines = new List<string>();
        List<IntPtr> open = found.FindAll(delegate(IntPtr hWnd) { return !watch.Done.Contains(hWnd); });
        if (open.Count == 0) {
            return lines;
        }
        // One pass under the lock, so that Settle returns only once every
        // notice up now has been answered and reported.
        lock (watch.Busy) {
            foreach (IntPtr hWnd in open) {
                if (current != null && !current()) {
                    break;
                }
                bool settled;
                string line = Answer(hWnd, folder, out settled);
                if (settled) {
                    watch.Done.Add(hWnd);
                }
                if (settled || watch.Reported.Add(hWnd)) {
                    lines.Add(line);
                    if (current != null) {
                        lock (gate) {
                            if (current()) { accepted.Add(line); }
                        }
                    }
                }
            }
        }
        return lines;
    }

    // Settled once it is pressed. A notice left alone is looked at again,
    // since its path may not have been there yet, and so is one that could
    // not be read, has no known button, or whose button could not be pressed.
    private static string Answer(IntPtr hWnd, string folder, out bool settled) {
        settled = false;
        string title = Text(hWnd);
        string action = "could not be read";
        List<string> texts = new List<string>();
        try {
            AutomationElement notice = AutomationElement.FromHandle(hWnd);
            foreach (AutomationElement element in notice.FindAll(TreeScope.Descendants, Condition.TrueCondition)) {
                AddText(texts, element.Current.Name);
                object pattern;
                if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern)) {
                    AddText(texts, ((ValuePattern)pattern).Current.Value);
                }
            }
            AutomationElement button = FindAcceptButton(notice);
            if (!String.IsNullOrEmpty(folder) && !NamesFileIn(texts, folder)) {
                action = "left alone, not about " + folder;
            }
            else if (button == null) {
                action = "no known button";
            }
            else {
                string name = button.Current.Name;
                action = "could not press " + name;
                ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                action = "pressed " + name;
                settled = true;
            }
        }
        catch (Exception error) {
            texts.Add("(" + error.GetType().Name + ": " + error.Message + ")");
        }
        return title + " | " + action + " | " + String.Join(" / ", texts.ToArray());
    }

    // Whether the notice's words name a path inside the folder, written in
    // its long or its short (8.3) form. A path the notice names is also
    // compared in its long form, which settles one whose folders are partly
    // short, as under a temp folder like C:\Users\RUNNER~1\AppData\Local\Temp.
    private static bool NamesFileIn(List<string> texts, string folder) {
        List<string> forms = new List<string>();
        forms.Add(folder);
        string longFolder = LongForm(folder);
        if (longFolder != null) { forms.Add(longFolder); }
        StringBuilder builder = new StringBuilder(1024);
        if (GetShortPathName(folder, builder, (uint)builder.Capacity) > 0) { forms.Add(builder.ToString()); }
        List<string> said = new List<string>(texts);
        foreach (string text in texts) {
            string longText = LongForm(text);
            if (longText != null) { said.Add(longText); }
        }
        foreach (string text in said) {
            foreach (string form in forms) {
                if (text.IndexOf(form.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase) >= 0) {
                    return true;
                }
            }
        }
        return false;
    }

    // The long form of a path, or of its folder when the file itself is not
    // there; null when it names no folder that exists.
    private static string LongForm(string path) {
        if (path.IndexOf(":\\", StringComparison.Ordinal) != 1 && !path.StartsWith("\\\\")) {
            return null;
        }
        StringBuilder builder = new StringBuilder(1024);
        if (GetLongPathName(path, builder, (uint)builder.Capacity) > 0) {
            return builder.ToString();
        }
        int slash = path.TrimEnd('\\').LastIndexOf('\\');
        if (slash <= 2) {
            return null;
        }
        builder = new StringBuilder(1024);
        return GetLongPathName(path.Substring(0, slash), builder, (uint)builder.Capacity) > 0
            ? builder.ToString() + path.Substring(slash)
            : null;
    }

    private static void AddText(List<string> texts, string text) {
        if (!String.IsNullOrWhiteSpace(text) && !texts.Contains(text.Trim())) {
            texts.Add(text.Trim());
        }
    }

    // Each Start is a run of its own: a thread from an earlier run that Stop
    // could not wait out keeps its own stop event, and publishes nothing.
    public static void Start(uint processId) {
        Stop();
        int run;
        lock (gate) {
            accepted.Clear();
            run = ++generation;
        }
        if (processId == 0) {
            return;
        }
        ManualResetEvent stop = new ManualResetEvent(false);
        Func<bool> current = delegate() { return Volatile.Read(ref generation) == run; };
        // The run's own lock: a thread of an earlier run still blocked in an
        // answer holds only its own, and Settle does not wait for it.
        Watch watch = new Watch();
        Thread thread = new Thread(delegate() {
            while (!stop.WaitOne(100) && current()) {
                try {
                    PollInto(watch, processId, null, current);
                }
                catch {
                }
            }
        });
        thread.IsBackground = true;
        lock (gate) { running = watch; }
        stopEvent = stop;
        watcherThread = thread;
        thread.Start();
    }

    public static bool Settle(int maxMilliseconds) {
        Watch watch;
        lock (gate) { watch = running; }
        if (watch == null) {
            return true;
        }
        if (!Monitor.TryEnter(watch.Busy, Math.Max(0, maxMilliseconds))) {
            return false;
        }
        Monitor.Exit(watch.Busy);
        return true;
    }

    public static void Stop() {
        lock (gate) {
            generation++;
            running = null;
        }
        if (stopEvent != null) {
            stopEvent.Set();
        }
        if (watcherThread != null && watcherThread.IsAlive) {
            watcherThread.Join(5000);
        }
        stopEvent = null;
        watcherThread = null;
    }

    public static string[] Accepted {
        get { lock (gate) { return accepted.ToArray(); } }
    }

    private static AutomationElement FindAcceptButton(AutomationElement notice) {
        AutomationElementCollection buttons = notice.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
        foreach (string wanted in AcceptButtons) {
            foreach (AutomationElement button in buttons) {
                string name = (button.Current.Name ?? "").Replace("&", "").Trim();
                if (String.Equals(name, wanted, StringComparison.OrdinalIgnoreCase)) {
                    return button;
                }
            }
        }
        return null;
    }

    private static List<IntPtr> FindNotices(uint processId) {
        List<IntPtr> found = new List<IntPtr>();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            uint owner;
            GetWindowThreadProcessId(hWnd, out owner);
            if ((processId == 0 || owner == processId) && IsWindowVisible(hWnd) && IsNoticeTitle(Text(hWnd))) {
                found.Add(hWnd);
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static string Text(IntPtr hWnd) {
        StringBuilder builder = new StringBuilder(512);
        GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }
}
"@
}
