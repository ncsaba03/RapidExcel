using System.Buffers;
using System.IO.MemoryMappedFiles;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace RapidExcel.Utils;

internal sealed class FastSharedStringReader : IDisposable
{
    private const int HEADER = 4;
    private const int INDEX_ENTRY_SIZE = 16;
    private readonly MemoryMappedFile _mmf;
    private readonly MemoryMappedViewAccessor _accessor;
    private readonly int _count;
    private readonly string _tempPath;
    private bool disposedValue;

    /// <summary>
    /// Initializes a new instance of the <see cref="FastSharedStringReader"/> class.
    /// </summary>
    /// <param name="tempPath">The temp path</param>
    /// <param name="mmf">MemoryMappedFile</param>
    /// <param name="accessor">MemoryPapedViewAccessor</param>
    /// <param name="count">Count of SSTs</param>
    private FastSharedStringReader(
        string tempPath,
        MemoryMappedFile mmf,
        MemoryMappedViewAccessor accessor,
        int count)
    {
        _tempPath = tempPath;
        _mmf = mmf;
        _accessor = accessor;
        _count = count;
    }

    /// <summary>
    /// Creates a FastSharedStringReader from the given SharedStringTablePart
    /// </summary>
    /// <param name="part"></param>
    /// <returns></returns>
    public static FastSharedStringReader Create(SharedStringTablePart part)
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"sst_{Guid.NewGuid()}.dat");
        MemoryMappedFile? mmf = null;
        MemoryMappedViewAccessor? accessor = null;

        try
        {
            int count = BuildIndexFile(part, tempPath);
            mmf = MemoryMappedFile.CreateFromFile(tempPath, FileMode.Open, null, 0, MemoryMappedFileAccess.Read);
            accessor = mmf.CreateViewAccessor(0, 0, MemoryMappedFileAccess.Read);

            return new FastSharedStringReader(tempPath, mmf, accessor, count);
        }
        catch
        {
            accessor?.Dispose();
            mmf?.Dispose();

            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); } catch { }
            }

            throw;
        }
    }

    public int Count => _count;

    /// <summary>
    /// Gets the string at the given index from the indexed shared string table
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public string GetString(int index)
    {
        ObjectDisposedException.ThrowIf(disposedValue, this);

        if (index < 0 || index >= _count)
            throw new ArgumentOutOfRangeException(nameof(index));

        long indexEntryPos = HEADER + ((long)INDEX_ENTRY_SIZE * index);
        long dataOffset = _accessor.ReadInt64(indexEntryPos);
        int length = _accessor.ReadInt32(indexEntryPos + 8);

        return ReadStringFromPointer(dataOffset, length);
    }

    /// <summary>
    /// Unsafe method to read string from memory-mapped file at given offset and length
    /// </summary>
    /// <param name="offset"></param>
    /// <param name="length"></param>
    /// <returns></returns>
    private unsafe string ReadStringFromPointer(long offset, int length)
    {
        byte* basePtr = null;
        _accessor.SafeMemoryMappedViewHandle.AcquirePointer(ref basePtr);
        try
        {
            var span = new ReadOnlySpan<byte>(basePtr + offset, length);
            return Encoding.UTF8.GetString(span);
        }
        finally
        {
            _accessor.SafeMemoryMappedViewHandle.ReleasePointer();
        }
    }

    /// <summary>
    /// Creates the index file for fast access
    /// <para>File Format:</para>
    /// <para>[Count (4)] [Index Table (12 + 4 padding * Count)] [Data Blob]</para>   
    /// </summary>
    /// <remarks>
    /// Using temporary files to avoid large memory consumption and two streams to avoid seeking back and forth. Then merge them into final file.   
    /// </remarks>
    /// <param name="part"></param>
    /// <param name="finalPath"></param>
    /// <returns></returns>
    private static int BuildIndexFile(SharedStringTablePart part, string finalPath)
    {
        using var dataStream = new FileStream(Path.GetTempFileName(), FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
        using var indexStream = new FileStream(Path.GetTempFileName(), FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);

        using var indexWriter = new BinaryWriter(indexStream);

        int count = 0;

        byte[]? buffer = ArrayPool<byte>.Shared.Rent(4096);
        var sb = new StringBuilder();

        try
        {
            using var partStream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = XmlReader.Create(partStream);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
                {

                    string text = ReadSharedStringItem(reader, sb);
                    int requiredSize = Encoding.UTF8.GetMaxByteCount(text.Length);

                    if (buffer.Length < requiredSize)
                    {
                        ArrayPool<byte>.Shared.Return(buffer);
                        buffer = ArrayPool<byte>.Shared.Rent(requiredSize);
                    }

                    int actualBytes = Encoding.UTF8.GetBytes(text, 0, text.Length, buffer, 0);
                    indexWriter.Write(dataStream.Position);
                    indexWriter.Write(actualBytes);

                    dataStream.Write(buffer.AsSpan(0, actualBytes));

                    count++;
                }
            }
        }
        finally
        {
            if (buffer != null)
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        using (var finalStream = new FileStream(finalPath, FileMode.Create))
        using (var finalWriter = new BinaryWriter(finalStream))
        {
            finalWriter.Write(count);
            indexStream.Position = 0;
            using (var indexReader = new BinaryReader(indexStream))
            {
                long indexSectionSize = (long)count * INDEX_ENTRY_SIZE;
                long baseDataOffset = HEADER + indexSectionSize;

                for (int i = 0; i < count; i++)
                {
                    long relativeOffset = indexReader.ReadInt64();
                    int length = indexReader.ReadInt32();
                    finalWriter.Write(baseDataOffset + relativeOffset);
                    finalWriter.Write(length);
                    finalWriter.Write(0); //Padding for alignment
                }
            }

            dataStream.Position = 0;
            dataStream.CopyTo(finalStream);
        }

        return count;
    }

    /// <summary>
    /// Reads a shared string item from the XmlReader
    /// </summary>
    /// <param name="reader"></param>
    /// <returns></returns>
    private static string ReadSharedStringItem(XmlReader reader, StringBuilder sb)
    {
        sb.Clear();
        while (!reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si")
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                sb.Append(reader.ReadElementContentAsString());
                continue;
            }

            reader.Read();
        }

        return sb.ToString();
    }

    /// <summary>
    /// Disposes the resources
    /// </summary>
    /// <param name="disposing"></param>
    private void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                _accessor?.Dispose();
                _mmf?.Dispose();
            }

            if (!string.IsNullOrEmpty(_tempPath) && File.Exists(_tempPath))
            {
                try { File.Delete(_tempPath); } catch { }
            }

            disposedValue = true;
        }
    }

    /// <summary>
    /// Disposes the resources
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Finalizer
    /// </summary>
    ~FastSharedStringReader()
    {
        Dispose(false);
    }
}