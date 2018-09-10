package com.unaware.poi.excel.streamreader;

import com.unaware.poi.excel.exception.ParameterException;
import com.unaware.poi.excel.exception.ReadException;
import com.unaware.poi.excel.ssimpl.StreamSheet;
import com.unaware.poi.excel.sstimpl.StreamSST;
import com.unaware.poi.excel.util.XmlUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.net.URI;
import java.nio.file.Files;
import java.security.GeneralSecurityException;
import java.util.*;

/**
 * @author Unaware
 * @date 2018/7/12 15:22
 * obtain the information of the workbook, sharedString, styles and sheets.
 * For workbook and styles, directly put all their contents in memory.
 * For sharedString, you can choose to put some or all of it in memory according to the parameter sstCacheSize.
 * For sheets, put some rows in memory, the number is up to the parameter rowCacheSize.
 */
public class WorkbookReader implements Iterable<Sheet>, AutoCloseable {
    /**
     * this holds the StreamSheet objects attached to this workbook
     */
    private final List<StreamSheet> sheets;
    private final List<Map<String, String>> sheetProperties;
    private final StreamReader.Builder builder;

    /**
     * shared string table - a cache of strings in this workbook
     */
    private SharedStringsTable sharedStringSource;

    /**
     * A collection of shared objects used for styling content,
     * e.g. fonts, cell styles, colors, etc.
     */
    private StylesTable stylesSource;

    private File tempFile;

    private OPCPackage OPCpkg;

    private File sstCache;

    private boolean use1904Dates = false;


    public WorkbookReader(StreamReader.Builder builder) {
        this.sheets = new ArrayList<>();
        this.sheetProperties = new ArrayList<>();
        this.builder = builder;
    }


    /**
     * init the inputStream, and create the temporary file for inputStream
     *
     * @param is
     */
    public void init(InputStream is) {
        try {
            if (builder.getBufferSize() <= 0) {
                throw new ParameterException("the bufferSize must be greater than 0");
            }
            tempFile = writeInputStreamToTempFile(is, builder.getBufferSize());
            //System.out.println("Debug: Created temp file [" + tempFile.getAbsolutePath() + "]");
            this.init(tempFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Extract the file and parse the file in XML format
     *
     * @param file
     */
    public void init(File file) {
        try {
            if (builder.getPassword() != null) {
                /*
                  Based on: https://poi.apache.org/encryption.html
                 */
                POIFSFileSystem poifs = new POIFSFileSystem(file);
                EncryptionInfo info = new EncryptionInfo(poifs);
                Decryptor d = Decryptor.getInstance(info);
                if (!d.verifyPassword(builder.getPassword())) {
                    throw new RuntimeException("Unable to process: this document is encrypted, the password is wrong!");
                }
                OPCpkg = OPCPackage.open(d.getDataStream(poifs));
            } else {
                OPCpkg = OPCPackage.open(file);
            }

            XSSFReader reader = new XSSFReader(OPCpkg);

            if (builder.getSstCacheSize() > 0) {
                sstCache = Files.createTempFile("", "").toFile();
                //System.out.println("Debug: Created sst cache file [" + sstCache.getAbsolutePath() + "]");
                sharedStringSource = StreamSST.getSharedStringTable(sstCache, builder.getSstCacheSize(), OPCpkg);
            } else {
                sharedStringSource = reader.getSharedStringsTable();
            }

            stylesSource = reader.getStylesTable();

            NodeList workbookPr = XmlUtils.searchForNodeList(XmlUtils.document(reader.getWorkbookData()), "/workbook/workbookPr");
            if (workbookPr.getLength() == 1) {
                final Node date1904 = workbookPr.item(0).getAttributes().getNamedItem("date1904");
                if (date1904 != null) {
                    use1904Dates = "1".equals(date1904.getTextContent());
                }
            }

            if (builder.getRowCacheSize() <= 0) {
                throw new ParameterException("the rowCacheSize must be greater than 0");
            }
            LoadSheets(reader, sharedStringSource, stylesSource, builder.getRowCacheSize());
        } catch (GeneralSecurityException e) {
            throw new ReadException("Unable to read workbook: Decryption failed", e);
        } catch (IOException e) {
            throw new ReadException("Unable to open workbook", e);
        } catch (OpenXML4JException | XMLStreamException e) {
            throw new ReadException("Unable to read workbook", e);
        }
    }

    /**
     * obtain the stream of all sheet from reader
     *
     * @param reader
     * @param sharedStringSource
     * @param stylesSource
     * @param rowCacheSize
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XMLStreamException
     */
    private void LoadSheets(XSSFReader reader, SharedStringsTable sharedStringSource, StylesTable stylesSource, int rowCacheSize) throws IOException, InvalidFormatException, XMLStreamException {
        /*
          obtain the name of all sheets
         */
        int numSheet = lookupSheetNames(reader);

        /*
          Some workbooks have multiple references to the same sheet.
          Need to filter them out before creating the XMLEventReader by keeping track of their URIs.
          The sheets are listed in order, so we must keep track of insertion order.
         */
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
        Map<URI, InputStream> sheetStreams = new LinkedHashMap<>();
        while (iter.hasNext()) {
            InputStream is = iter.next();
            sheetStreams.put(iter.getSheetPart().getPartName().getURI(), is);
        }

        /*
          Iterate over the loaded streams
         */
        int i = 0;
        for (URI uri : sheetStreams.keySet()) {
            //InputStream mergedStream = reader.getSheet("rId" + (i+1));
            if (builder.getSheetIndex() != -1) {
                if (builder.getSheetIndex() < 0 || builder.getSheetIndex() >= numSheet) {
                    throw new ParameterException("the sheet with the index '" + builder.getSheetIndex() + "' does not exist");
                }
                if (builder.getSheetIndex() == i) {
                    //XMLEventReader parser = StaxHelper.newXMLInputFactory().createXMLEventReader(sheetStreams.get(uri));
                    sheets.add(new StreamSheet(sheetProperties.get(i).get("name"),
                            new SheetReader(sharedStringSource, stylesSource, sheetStreams.get(uri), reader.getSheet("rId" + (i+1)), use1904Dates, rowCacheSize)));
                    break;
                }
            } else {
                //XMLEventReader parser = StaxHelper.newXMLInputFactory().createXMLEventReader(sheetStreams.get(uri));
                sheets.add(new StreamSheet(sheetProperties.get(i).get("name"),
                        new SheetReader(sharedStringSource, stylesSource, sheetStreams.get(uri), reader.getSheet("rId" + (i+1)), use1904Dates, rowCacheSize)));
            }
            i++;
        }
    }

    /**
     * obtain the name of all sheets
     *
     * @param reader
     * @return the number of sheets
     * @throws IOException
     * @throws InvalidFormatException
     */
    private int lookupSheetNames(XSSFReader reader) throws IOException, InvalidFormatException {
        sheetProperties.clear();
        NodeList nodeList = XmlUtils.searchForNodeList(XmlUtils.document(reader.getWorkbookData()), "/workbook/sheets/sheet");
        for (int i = 0; i < nodeList.getLength(); ++i) {
            Map<String, String> props = new HashMap<>();
            props.put("name", nodeList.item(i).getAttributes().getNamedItem("name").getTextContent());

            Node state = nodeList.item(i).getAttributes().getNamedItem("state");
            props.put("state", state == null ? "visible" : state.getTextContent());
            sheetProperties.add(props);
        }
        return nodeList.getLength();
    }

    /**
     * create the temporary file
     *
     * @param is
     * @param bufferSize
     * @return
     * @throws IOException
     */
    private static File writeInputStreamToTempFile(InputStream is, int bufferSize) throws IOException {
        File file = Files.createTempFile("temp_", ".xlsx").toFile();
        try (FileOutputStream fos = new FileOutputStream(file)) {
            int num;
            byte[] bytes = new byte[bufferSize];
            while ((num = is.read(bytes)) != -1) {
                fos.write(bytes, 0, num);
            }
            is.close();
            return file;
        }
    }

    /**
     * Return a object representing a collection of shared objects used for styling content,
     * e.g. fonts, cell styles, colors, etc.
     */
    public StylesTable getStylesSource() {
        return this.stylesSource;
    }

    public List<StreamSheet> getSheets() {
        return sheets;
    }

    public List<Map<String, String>> getSheetProperties() {
        return sheetProperties;
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
        try {
            for (StreamSheet sheet : sheets) {
                sheet.getReader().close();
            }
            OPCpkg.revert();
        } finally {
            if (tempFile != null) {
                //System.out.println("Debug: Deleting tmp file [" + tempFile.getAbsolutePath() + "]");
                tempFile.delete();
            }
            if (sharedStringSource instanceof StreamSST) {
                //System.out.println("Debug: Deleting sst cache file [" + this.sstCache.getAbsolutePath() + "]");
                ((StreamSST) sharedStringSource).close();
                sstCache.delete();
            }
            sheetProperties.clear();
        }
    }

    @Override
    public Iterator<Sheet> iterator() {
        return new StreamSheetIterator(sheets.iterator());
    }

    /**
     * add inner class to achieve iterator Sheets
     * implements Iterator<Sheet>
     */
    private static class StreamSheetIterator implements Iterator<Sheet> {
        private final Iterator<StreamSheet> iterator;

        public StreamSheetIterator(Iterator<StreamSheet> iterator) {
            this.iterator = iterator;
        }

        @Override
        public boolean hasNext() {
            return iterator.hasNext();
        }

        @Override
        public Sheet next() {
            return iterator.next();
        }

        @Override
        public void remove() {
            throw new RuntimeException("NotSupported");
        }
    }
}
