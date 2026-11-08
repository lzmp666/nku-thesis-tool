using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class UndoManagerTests
{
    [Fact]
    public void UndoRestoresFileToStateBeforeModify()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var path = Path.Combine(dir, "thesis.docx");
            File.WriteAllText(path, "original");

            undo.BackupBeforeModify(path, changeDescription: "测试批量修改");
            File.WriteAllText(path, "modified-and-wrong");

            undo.Undo(path);

            Assert.Equal("original", File.ReadAllText(path));
            // 撤销后该备份记录被移除
            Assert.Empty(undo.ListBackups());
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void UndoWithoutBackupThrowsReadableError()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var path = Path.Combine(dir, "thesis.docx");
            File.WriteAllText(path, "original");

            var ex = Assert.Throws<InvalidOperationException>(() => undo.Undo(path));
            Assert.Contains("没有找到备份", ex.Message);
            Assert.Equal("original", File.ReadAllText(path));
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void RepeatedUndoRollsBackToEachPreviousSnapshot()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var path = Path.Combine(dir, "thesis.docx");
            File.WriteAllText(path, "v0");

            undo.BackupBeforeModify(path);
            File.WriteAllText(path, "v1");

            // 第二次批量修改前再快照
            Thread.Sleep(10);
            undo.BackupBeforeModify(path);
            File.WriteAllText(path, "v2");

            undo.Undo(path);
            Assert.Equal("v1", File.ReadAllText(path));

            undo.Undo(path);
            Assert.Equal("v0", File.ReadAllText(path));

            // 备份记录已用尽
            Assert.Empty(undo.ListBackups());
            Assert.Throws<InvalidOperationException>(() => undo.Undo(path));
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void ManifestRecordsSourceOutputAndChangeDescription()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var source = Path.Combine(dir, "thesis.docx");
            var output = Path.Combine(dir, "thesis_nku.docx");
            File.WriteAllText(source, "original");

            undo.BackupBeforeModify(source, output, "套用南开格式：正文宋体小四");

            var record = Assert.Single(undo.ListBackups());
            Assert.Equal(Path.GetFullPath(source), record.SourcePath);
            Assert.Equal(Path.GetFullPath(output), record.OutputPath);
            Assert.Equal("套用南开格式：正文宋体小四", record.ChangeDescription);
            Assert.False(string.IsNullOrWhiteSpace(record.BackupId));
            Assert.False(string.IsNullOrWhiteSpace(record.CreatedAt));

            // 备份文件确实存在
            var backupPath = Path.Combine(backupDir, record.BackupFileName);
            Assert.True(File.Exists(backupPath));
            Assert.Equal("original", File.ReadAllText(backupPath));
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void UndoOnOutputPathRestoresContentEvenWhenOutputDiffersFromSource()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var source = Path.Combine(dir, "thesis.docx");
            var output = Path.Combine(dir, "thesis_nku.docx");
            File.WriteAllText(source, "original");

            undo.BackupBeforeModify(source, output, "批量修改");
            File.WriteAllText(output, "modified-and-wrong");

            // 用户对输出文件一键撤销
            undo.Undo(output);

            Assert.Equal("original", File.ReadAllText(output));
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void CleanOldBackupsKeepsOnlyNewestRecords()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var path = Path.Combine(dir, "thesis.docx");
            File.WriteAllText(path, "v0");

            for (var i = 0; i < 5; i++)
            {
                undo.BackupBeforeModify(path);
                Thread.Sleep(10);
                File.WriteAllText(path, $"v{i + 1}");
            }

            undo.CleanOldBackups(keepCount: 2);

            Assert.Equal(2, undo.ListBackups().Count);
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void BackupBeforeModifyThrowsWhenFileMissing()
    {
        var undo = new UndoManager(NewTempDir());
        Assert.Throws<FileNotFoundException>(() => undo.BackupBeforeModify("Z:\\not-exist.docx"));
    }

    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "undo-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void TryDelete(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch (IOException) { }
    }
}
