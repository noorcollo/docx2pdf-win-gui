using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;
using System.IO;

namespace Docx2PdfGui;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        WriteLog($"Launch attempt · {Environment.OSVersion} · {Environment.Is64BitProcess}-bit process");
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) => ShowStartupError(e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex) ShowStartupError(ex);
        };
        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            ShowStartupError(ex);
        }
    }

    private static void ShowStartupError(Exception exception)
    {
        var logPath = GetLogPath();
        WriteLog(exception.ToString());
        try
        {
            MessageBox.Show($"DOCX to PDF could not start.\n\n{exception.GetType().Name}: {exception.Message}\n\nDetails were saved to:\n{logPath}",
                "DOCX to PDF — startup error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch { }
    }

    private static string GetLogPath() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Docx2Pdf", "startup.log");

    private static void WriteLog(string message)
    {
        try
        {
            var path = GetLogPath();
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n\r\n");
        }
        catch { }
    }
}

internal sealed class SourceItem
{
    public required string FullPath { get; init; }
    public required string RelativePdfPath { get; init; }
    public string Status { get; set; } = "Ready";
}

internal sealed record ConversionResult(string Path, bool Success, string? Error);

internal sealed class MainForm : Form
{
    private readonly Color ink = Color.FromArgb(24, 39, 65);
    private readonly Color muted = Color.FromArgb(111, 126, 148);
    private readonly Color blue = Color.FromArgb(54, 101, 238);
    private readonly Color green = Color.FromArgb(24, 157, 120);
    private readonly List<SourceItem> queue = new();
    private ListView list = null!;
    private TextBox outputPath = null!;
    private CheckBox recurse = null!;
    private ProgressBar progress = null!;
    private Label statusLabel = null!;
    private Label countLabel = null!;
    private Button convertButton = null!;
    private Button addFilesButton = null!;
    private Button addFolderButton = null!;
    private Button removeButton = null!;
    private Button clearButton = null!;
    private Panel dropPanel = null!;
    private bool isConverting;
    private volatile bool cancelRequested;

    public MainForm()
    {
        Text = "DOCX to PDF";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(860, 640);
        Size = new Size(1040, 760);
        BackColor = Color.FromArgb(245, 247, 251);
        ForeColor = ink;
        Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);
        AllowDrop = true;
        BuildUi();
        DragEnter += HandleDragEnter;
        DragDrop += HandleDragDrop;
        UpdateCounts();
    }

    private void BuildUi()
    {
        var page = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(26, 22, 26, 18),
            BackColor = BackColor
        };
        page.RowStyles.Add(new RowStyle(SizeType.Absolute, 88));
        page.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        page.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
        Controls.Add(page);

        var header = new Panel { Dock = DockStyle.Fill, BackColor = BackColor };
        var mark = new Panel
        {
            Location = new Point(0, 8), Size = new Size(54, 54), BackColor = blue
        };
        mark.Paint += (_, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = RoundedRect(new Rectangle(0, 0, mark.Width - 1, mark.Height - 1), 13);
            using var brush = new SolidBrush(blue);
            e.Graphics.FillPath(brush, path);
            using var f = new Font("Segoe UI", 10, FontStyle.Bold);
            TextRenderer.DrawText(e.Graphics, "W→P", f, new Rectangle(2, 4, mark.Width - 4, mark.Height - 8), Color.White,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        };
        header.Controls.Add(mark);
        header.Controls.Add(new Label
        {
            Text = "DOCX to PDF", Font = new Font("Segoe UI Semibold", 20, FontStyle.Bold),
            ForeColor = ink, Location = new Point(72, 7), AutoSize = true
        });
        header.Controls.Add(new Label
        {
            Text = "Batch convert with Microsoft Word’s original PDF engine",
            Font = new Font("Segoe UI", 10), ForeColor = muted, Location = new Point(74, 43), AutoSize = true
        });
        var wordBadge = new Label
        {
            Text = "POWERED BY WORD", AutoSize = false, TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 8, FontStyle.Bold), ForeColor = Color.FromArgb(43, 91, 77),
            BackColor = Color.FromArgb(225, 244, 237), Size = new Size(138, 30), Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Location = new Point(Width - 235, 20)
        };
        header.Controls.Add(wordBadge);
        header.Resize += (_, _) => wordBadge.Left = header.ClientSize.Width - wordBadge.Width - 4;
        page.Controls.Add(header, 0, 0);

        var body = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, BackColor = BackColor,
            Margin = new Padding(0)
        };
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
        body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 90));
        page.Controls.Add(body, 0, 1);

        var actions = new Panel { Dock = DockStyle.Fill, BackColor = BackColor };
        addFilesButton = MakeButton("＋  Add Word files", blue, Color.White, 154, 42, true);
        addFilesButton.Location = new Point(0, 8);
        addFilesButton.Click += (_, _) => ChooseFiles();
        addFolderButton = MakeButton("▣  Add folder", Color.White, ink, 132, 42, false);
        addFolderButton.Location = new Point(164, 8);
        addFolderButton.Click += (_, _) => ChooseFolder();
        removeButton = MakeButton("Remove selected", Color.White, muted, 130, 42, false);
        removeButton.Location = new Point(310, 8);
        removeButton.Click += (_, _) => RemoveSelected();
        clearButton = MakeButton("Clear list", Color.White, muted, 92, 42, false);
        clearButton.Location = new Point(452, 8);
        clearButton.Click += (_, _) => ClearQueue();
        actions.Controls.AddRange(new Control[] { addFilesButton, addFolderButton, removeButton, clearButton });
        body.Controls.Add(actions, 0, 0);

        var queueCard = new Panel
        {
            Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(16),
            Margin = new Padding(0, 0, 0, 13)
        };
        queueCard.Paint += PaintCardBorder;
        var queueLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = Color.White };
        queueLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        queueLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        queueCard.Controls.Add(queueLayout);
        var queueHeader = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
        queueHeader.Controls.Add(new Label
        {
            Text = "Conversion queue", Font = new Font("Segoe UI Semibold", 11, FontStyle.Bold), ForeColor = ink,
            Location = new Point(0, 3), AutoSize = true
        });
        countLabel = new Label
        {
            Text = "0 documents", Font = new Font("Segoe UI", 9), ForeColor = muted,
            Dock = DockStyle.Right, TextAlign = ContentAlignment.MiddleRight, Width = 170
        };
        queueHeader.Controls.Add(countLabel);
        queueLayout.Controls.Add(queueHeader, 0, 0);

        list = new ListView();
        list.Dock = DockStyle.Fill;
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = true;
        list.BorderStyle = BorderStyle.None;
        list.BackColor = Color.White;
        list.ForeColor = ink;
        list.Font = new Font("Segoe UI", 9.5F);
        list.ShowItemToolTips = true;
        list.Columns.Add("FILE", 285);
        list.Columns.Add("LOCATION", 380);
        list.Columns.Add("STATUS", 120);
        list.ItemSelectionChanged += (_, _) => removeButton.Enabled = !isConverting && list.SelectedItems.Count > 0;
        queueLayout.Controls.Add(list, 0, 1);
        body.Controls.Add(queueCard, 0, 1);

        dropPanel = new Panel
        {
            Dock = DockStyle.Fill, BackColor = Color.FromArgb(249, 251, 255),
            Margin = new Padding(0, 0, 0, 14), AllowDrop = true
        };
        dropPanel.Paint += PaintDropPanel;
        dropPanel.DragEnter += HandleDragEnter;
        dropPanel.DragDrop += HandleDragDrop;
        dropPanel.Controls.Add(new Label
        {
            Text = "Drop DOCX files or folders here", AutoSize = false, Dock = DockStyle.Top, Height = 28,
            TextAlign = ContentAlignment.BottomCenter, Font = new Font("Segoe UI Semibold", 10, FontStyle.Bold),
            ForeColor = ink, BackColor = Color.Transparent
        });
        dropPanel.Controls.Add(new Label
        {
            Text = "Folders are scanned recursively · Word lock files are ignored", AutoSize = false, Dock = DockStyle.Bottom,
            Height = 25, TextAlign = ContentAlignment.TopCenter, Font = new Font("Segoe UI", 8.5F),
            ForeColor = muted, BackColor = Color.Transparent
        });
        body.Controls.Add(dropPanel, 0, 2);

        var outputPanel = new Panel { Dock = DockStyle.Fill, BackColor = BackColor };
        outputPanel.Controls.Add(new Label
        {
            Text = "Save PDFs to", Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold), ForeColor = ink,
            Location = new Point(0, 2), AutoSize = true
        });
        outputPath = new TextBox();
        outputPath.ReadOnly = true;
        outputPath.PlaceholderText = "Same folder as each Word file";
        outputPath.Location = new Point(0, 28);
        outputPath.Size = new Size(550, 30);
        outputPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        outputPath.BackColor = Color.White;
        outputPath.BorderStyle = BorderStyle.FixedSingle;
        outputPanel.Controls.Add(outputPath);
        var browse = MakeButton("Browse…", Color.White, ink, 100, 31, false);
        browse.Location = new Point(560, 27);
        browse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        browse.Click += (_, _) => ChooseOutput();
        outputPanel.Controls.Add(browse);
        recurse = new CheckBox();
        recurse.Text = "Include subfolders";
        recurse.Checked = true;
        recurse.AutoSize = true;
        recurse.ForeColor = muted;
        recurse.Font = new Font("Segoe UI", 8.5F);
        recurse.Location = new Point(0, 62);
        recurse.CheckedChanged += (_, _) => { };
        outputPanel.Controls.Add(recurse);
        var clearOutput = new LinkLabel
        {
            Text = "Reset", AutoSize = true, Location = new Point(112, 64), LinkColor = blue,
            ActiveLinkColor = blue, VisitedLinkColor = blue, Font = new Font("Segoe UI", 8.5F)
        };
        clearOutput.Click += (_, _) => outputPath.Clear();
        outputPanel.Controls.Add(clearOutput);
        outputPanel.Resize += (_, _) =>
        {
            browse.Left = outputPanel.ClientSize.Width - browse.Width;
            outputPath.Width = Math.Max(100, browse.Left - 10);
        };
        body.Controls.Add(outputPanel, 0, 3);

        var footer = new Panel { Dock = DockStyle.Fill, BackColor = BackColor };
        progress = new ProgressBar { Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100, Value = 0,
            Location = new Point(0, 17), Size = new Size(260, 10), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };
        statusLabel = new Label { Text = "Ready when you are", ForeColor = muted, Font = new Font("Segoe UI", 8.5F),
            Location = new Point(0, 0), AutoSize = true };
        footer.Controls.Add(statusLabel);
        footer.Controls.Add(progress);
        convertButton = MakeButton("Convert to PDF  →", blue, Color.White, 200, 44, true);
        convertButton.Location = new Point(footer.ClientSize.Width - 200, 5);
        convertButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        convertButton.Click += ConvertButtonClick;
        footer.Controls.Add(convertButton);
        footer.Resize += (_, _) => progress.Width = Math.Max(80, convertButton.Left - 20);
        page.Controls.Add(footer, 0, 2);
    }

    private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
    {
        var d = radius * 2;
        var path = new GraphicsPath();
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    private static Button MakeButton(string text, Color back, Color fore, int width, int height, bool primary)
    {
        var b = new Button
        {
            Text = text, Size = new Size(width, height), BackColor = back, ForeColor = fore,
            FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI Semibold", 9, FontStyle.Bold),
            Cursor = Cursors.Hand, UseVisualStyleBackColor = false, TabStop = true
        };
        b.FlatAppearance.BorderColor = primary ? back : Color.FromArgb(220, 226, 235);
        b.FlatAppearance.BorderSize = primary ? 0 : 1;
        b.FlatAppearance.MouseOverBackColor = primary ? Color.FromArgb(42, 84, 210) : Color.FromArgb(242, 245, 250);
        return b;
    }

    private static void PaintCardBorder(object? sender, PaintEventArgs e)
    {
        if (sender is not Control c) return;
        using var pen = new Pen(Color.FromArgb(226, 232, 241));
        e.Graphics.DrawRectangle(pen, 0, 0, c.Width - 1, c.Height - 1);
    }

    private static void PaintDropPanel(object? sender, PaintEventArgs e)
    {
        if (sender is not Control c) return;
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = new Pen(Color.FromArgb(176, 194, 235), 1.4F) { DashPattern = new[] { 4F, 4F } };
        using var path = RoundedRect(new Rectangle(1, 1, c.Width - 3, c.Height - 3), 9);
        e.Graphics.DrawPath(pen, path);
    }

    private void ChooseFiles()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Select Word documents", Filter = "Word documents (*.docx)|*.docx",
            Multiselect = true, CheckFileExists = true
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            foreach (var file in dialog.FileNames) AddSource(file, Path.GetFileName(file));
    }

    private void ChooseFolder()
    {
        using var dialog = new FolderBrowserDialog { Description = "Choose a folder containing DOCX documents", UseDescriptionForTitle = true };
        if (dialog.ShowDialog(this) == DialogResult.OK) AddFolder(dialog.SelectedPath);
    }

    private void ChooseOutput()
    {
        using var dialog = new FolderBrowserDialog { Description = "Choose where converted PDFs should be saved", UseDescriptionForTitle = true };
        if (Directory.Exists(outputPath.Text)) dialog.SelectedPath = outputPath.Text;
        if (dialog.ShowDialog(this) == DialogResult.OK) outputPath.Text = dialog.SelectedPath;
    }

    private void AddFolder(string folder)
    {
        try
        {
            var rootName = new DirectoryInfo(folder).Name;
            var option = recurse.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.EnumerateFiles(folder, "*.docx", option)
                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList();
            foreach (var file in files)
            {
                var relative = Path.GetRelativePath(folder, file);
                AddSource(file, Path.Combine(rootName, relative));
            }
            if (files.Count == 0) SetStatus("No DOCX documents found in that folder.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not read that folder.\n\n{ex.Message}", "Folder access error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void AddDropped(string[] paths)
    {
        foreach (var path in paths)
        {
            if (Directory.Exists(path)) AddFolder(path);
            else if (File.Exists(path) && Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase)
                     && !Path.GetFileName(path).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                AddSource(path, Path.GetFileName(path));
        }
    }

    private void HandleDragEnter(object? sender, DragEventArgs e)
    {
        e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true ? DragDropEffects.Copy : DragDropEffects.None;
    }

    private void HandleDragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] paths) AddDropped(paths);
    }

    private void AddSource(string path, string relative)
    {
        var full = Path.GetFullPath(path);
        if (queue.Any(q => q.FullPath.Equals(full, StringComparison.OrdinalIgnoreCase))) return;
        var relPdf = Path.ChangeExtension(relative, ".pdf");
        var reserved = new HashSet<string>(queue.Select(q => q.RelativePdfPath), StringComparer.OrdinalIgnoreCase);
        if (reserved.Contains(relPdf))
        {
            var directory = Path.GetDirectoryName(relative) ?? "";
            var stem = Path.GetFileNameWithoutExtension(relative);
            var ext = Path.GetExtension(relative);
            var n = 2;
            do { relPdf = Path.ChangeExtension(Path.Combine(directory, $"{stem} ({n++}){ext}"), ".pdf"); }
            while (reserved.Contains(relPdf));
        }
        var item = new SourceItem { FullPath = full, RelativePdfPath = relPdf };
        queue.Add(item);
        var row = new ListViewItem(Path.GetFileName(full));
        row.SubItems.Add(Path.GetDirectoryName(full) ?? "");
        row.SubItems.Add(item.Status);
        row.Tag = item;
        row.ToolTipText = full;
        list.Items.Add(row);
        UpdateCounts();
    }

    private void RemoveSelected()
    {
        if (isConverting) return;
        foreach (ListViewItem row in list.SelectedItems)
        {
            if (row.Tag is SourceItem source) queue.Remove(source);
            list.Items.Remove(row);
        }
        UpdateCounts();
    }

    private void ClearQueue()
    {
        if (isConverting) return;
        queue.Clear();
        list.Items.Clear();
        progress.Value = 0;
        UpdateCounts();
        SetStatus("Queue cleared.");
    }

    private void UpdateCounts()
    {
        countLabel.Text = $"{queue.Count} document{(queue.Count == 1 ? "" : "s")}";
        convertButton.Enabled = !isConverting && queue.Count > 0;
        removeButton.Enabled = !isConverting && list.SelectedItems.Count > 0;
        clearButton.Enabled = !isConverting && queue.Count > 0;
        addFilesButton.Enabled = addFolderButton.Enabled = !isConverting;
    }

    private void SetStatus(string text, Color? color = null)
    {
        statusLabel.Text = text;
        statusLabel.ForeColor = color ?? muted;
    }

    private async Task StartConversion()
    {
        if (isConverting || queue.Count == 0) return;
        var outputRoot = outputPath.Text.Trim();
        if (outputRoot.Length > 0)
        {
            try { Directory.CreateDirectory(outputRoot); }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Cannot use the selected output folder.\n\n{ex.Message}", "Output folder error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        if (Type.GetTypeFromProgID("Word.Application") is null)
        {
            MessageBox.Show(this, "Microsoft Word was not found. Install the desktop version of Microsoft Word (2013 or later) and try again.",
                "Microsoft Word required", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var existingPdfs = queue.Select(source => outputRoot.Length == 0
                ? Path.ChangeExtension(source.FullPath, ".pdf")
                : Path.Combine(outputRoot, source.RelativePdfPath))
            .Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (existingPdfs.Count > 0)
        {
            var preview = string.Join(Environment.NewLine, existingPdfs.Take(5).Select(Path.GetFileName));
            var more = existingPdfs.Count > 5 ? $"{Environment.NewLine}…and {existingPdfs.Count - 5} more" : "";
            var answer = MessageBox.Show(this,
                $"{existingPdfs.Count} PDF file(s) already exist and will be replaced:\n\n{preview}{more}\n\nContinue?",
                "Confirm PDF overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) return;
        }

        isConverting = true;
        cancelRequested = false;
        UpdateCounts();
        convertButton.Text = "Cancel conversion";
        convertButton.Enabled = true;
        convertButton.BackColor = Color.FromArgb(201, 70, 77);
        convertButton.FlatAppearance.BorderColor = convertButton.BackColor;
        // The same button handler changes to a safe, cooperative cancel action while running.
        progress.Value = 0;
        SetStatus("Starting Microsoft Word…");
        var snapshot = queue.ToArray();
        var results = new List<ConversionResult>();
        var done = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            object? word = null;
            try
            {
                var wordType = Type.GetTypeFromProgID("Word.Application") ?? throw new InvalidOperationException("Microsoft Word is not installed.");
                word = Activator.CreateInstance(wordType) ?? throw new InvalidOperationException("Microsoft Word could not be started.");
                dynamic app = word;
                app.Visible = false;
                app.DisplayAlerts = 0;
                for (int i = 0; i < snapshot.Length; i++)
                {
                    if (cancelRequested) break;
                    var source = snapshot[i];
                    var target = outputRoot.Length == 0
                        ? Path.ChangeExtension(source.FullPath, ".pdf")
                        : Path.Combine(outputRoot, source.RelativePdfPath);
                    try
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                        if (File.Exists(target)) File.Delete(target);
                        this.BeginInvoke(() => SetRowStatus(source, "Converting…", null));
                        object? doc = null;
                        try
                        {
                            doc = app.Documents.Open(source.FullPath, false, true);
                            dynamic document = doc;
                            document.ExportAsFixedFormat(target, 17);
                            document.Close(false);
                            results.Add(new ConversionResult(source.FullPath, true, null));
                        }
                        finally
                        {
                            if (doc is not null && Marshal.IsComObject(doc)) Marshal.FinalReleaseComObject(doc);
                        }
                        this.BeginInvoke(() => SetRowStatus(source, "Done", null));
                    }
                    catch (Exception ex)
                    {
                        var message = ex.GetBaseException().Message;
                        results.Add(new ConversionResult(source.FullPath, false, message));
                        this.BeginInvoke(() => SetRowStatus(source, "Failed", message));
                    }
                    var completed = i + 1;
                    this.BeginInvoke(() =>
                    {
                        progress.Value = Math.Min(100, (int)Math.Round(completed * 100.0 / snapshot.Length));
                        SetStatus($"Processed {completed} of {snapshot.Length} documents…");
                    });
                }
            }
            catch (Exception ex)
            {
                this.BeginInvoke(() => MessageBox.Show(this, "Microsoft Word could not be started. Make sure desktop Word is installed and activated.\n\n" + ex.GetBaseException().Message,
                    "Word startup error", MessageBoxButtons.OK, MessageBoxIcon.Error));
            }
            finally
            {
                if (word is not null)
                {
                    try { ((dynamic)word).Quit(false); } catch { }
                    if (Marshal.IsComObject(word)) Marshal.FinalReleaseComObject(word);
                }
                done.TrySetResult();
            }
        }) { IsBackground = true, Name = "Word DOCX conversion" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        await done.Task;
        isConverting = false;
        convertButton.Text = "Convert to PDF  →";
        convertButton.BackColor = blue;
        convertButton.FlatAppearance.BorderColor = blue;
        UpdateCounts();
        var succeeded = results.Count(r => r.Success);
        var failed = results.Count(r => !r.Success);
        var wasCancelled = cancelRequested;
        SetStatus(wasCancelled
            ? $"Stopped · {succeeded} converted, {failed} failed"
            : $"Finished · {succeeded} converted, {failed} failed", failed > 0 ? Color.FromArgb(190, 77, 82) : green);
        if (failed > 0)
        {
            var firstError = results.First(r => !r.Success);
            var shortError = firstError.Error ?? "Unknown conversion error.";
            MessageBox.Show(this, $"Completed with {failed} failed document{(failed == 1 ? "" : "s")}.\n\nExample: {Path.GetFileName(firstError.Path)}\n{shortError}",
                "Conversion finished", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void SetRowStatus(SourceItem source, string status, string? error)
    {
        source.Status = status;
        foreach (ListViewItem row in list.Items)
        {
            if (ReferenceEquals(row.Tag, source))
            {
                row.SubItems[2].Text = status;
                row.ToolTipText = error is null ? source.FullPath : source.FullPath + "\n\n" + error;
                break;
            }
        }
    }

    private void CancelConversion(object? sender, EventArgs e)
    {
        cancelRequested = true;
        convertButton.Enabled = false;
        SetStatus("Finishing the current document before stopping…");
    }

    private async void ConvertButtonClick(object? sender, EventArgs e)
    {
        if (isConverting) CancelConversion(sender, e);
        else await StartConversion();
    }
}
