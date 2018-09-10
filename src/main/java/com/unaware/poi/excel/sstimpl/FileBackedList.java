package com.unaware.poi.excel.sstimpl;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Unaware
 * @Description: ${description}
 * @Title: FileBackedList
 * @ProjectName excel
 * @date 2018/7/13 17:05
 */
public class FileBackedList<T> implements AutoCloseable {
    private final static ObjectMapper mapper;
    private final Class<T> type;
    private final List<Long> pointers;
    private final RandomAccessFile raFile;
    private final FileChannel channel;
    private final Map<Integer, T> cache;

    private long fileSize;

    static {
        mapper = new ObjectMapper().setSerializationInclusion(JsonInclude.Include.NON_NULL);
    }

    public FileBackedList(Class<T> type, File file, final int cacheSize) throws IOException {
        this.type = type;
        this.raFile = new RandomAccessFile(file, "rw");
        this.channel = raFile.getChannel();
        this.fileSize = raFile.length();
        this.cache = new LinkedHashMap<Integer, T>(cacheSize, 0.75f, true) {
            public boolean removeEldestEntry(Map.Entry eldest) {
                return size() > cacheSize;
            }
        };
        pointers = new ArrayList<>();
    }

    public void add(T obj) {
        try {
            writeToFile(obj);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public T getAt(int index) {
        if (cache.containsKey(index)) {
            return cache.get(index);
        }

        try {
            T val = readFromFile(pointers.get(index));
            cache.put(index, val);
            return val;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void writeToFile(T obj) throws IOException {
        synchronized (channel) {
            ByteBuffer bytes = ByteBuffer.wrap(mapper.writeValueAsBytes(obj));
            ByteBuffer length = ByteBuffer.allocate(4).putInt(bytes.array().length);

            channel.position(fileSize);
            pointers.add(channel.position());
            length.flip();
            channel.write(length);
            channel.write(bytes);

            fileSize += 4 + bytes.array().length;
        }
    }

    private T readFromFile(long pointer) throws IOException {
        synchronized (channel) {
            FileChannel fc = channel.position(pointer);

            //get length of entry
            ByteBuffer buffer = ByteBuffer.wrap(new byte[4]);
            fc.read(buffer);
            buffer.flip();
            int length = buffer.getInt();

            //read entry
            buffer = ByteBuffer.wrap(new byte[length]);
            fc.read(buffer);
            buffer.flip();

            return mapper.readValue(buffer.array(), type);
        }
    }

    /**
     * Closes this resource, relinquishing any underlying resources.
     * This method is invoked automatically on objects managed by the
     * {@code try}-with-resources statement.
     * <p>
     * <p>While this interface method is declared to throw {@code
     * exception}, implementers are <em>strongly</em> encouraged to
     * declare concrete implementations of the {@code close} method to
     * throw more specific exceptions, or to throw no exception at all
     * if the close operation cannot fail.
     * <p>
     * <p> Cases where the close operation may fail require careful
     * attention by implementers. It is strongly advised to relinquish
     * the underlying resources and to internally <em>mark</em> the
     * resource as closed, prior to throwing the exception. The {@code
     * close} method is unlikely to be invoked more than once and so
     * this ensures that the resources are released in a timely manner.
     * Furthermore it reduces problems that could arise when the resource
     * wraps, or is wrapped, by another resource.
     * <p>
     * <p><em>Implementers of this interface are also strongly advised
     * to not have the {@code close} method throw {@link
     * InterruptedException}.</em>
     * <p>
     * This exception interacts with a thread's interrupted status,
     * and runtime misbehavior is likely to occur if an {@code
     * InterruptedException} is {@linkplain Throwable#addSuppressed
     * suppressed}.
     * <p>
     * More generally, if it would cause problems for an
     * exception to be suppressed, the {@code AutoCloseable.close}
     * method should not throw it.
     * <p>
     * <p>Note that unlike the {@link Closeable#close close}
     * method of {@link Closeable}, this {@code close} method
     * is <em>not</em> required to be idempotent.  In other words,
     * calling this {@code close} method more than once may have some
     * visible side effect, unlike {@code Closeable.close} which is
     * required to have no effect if called more than once.
     * <p>
     * However, implementers of this interface are strongly encouraged
     * to make their {@code close} methods idempotent.
     *
     * @throws Exception if this resource cannot be closed
     */
    @Override
    public void close() throws Exception {
        raFile.close();
    }
}
