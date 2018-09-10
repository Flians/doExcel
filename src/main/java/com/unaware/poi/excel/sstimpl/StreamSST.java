package com.unaware.poi.excel.sstimpl;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;
import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @author Unaware
 * @Description: parse the SharedString.xml
 * @Title: StreamSST
 * @ProjectName excel
 * @date 2018/7/13 13:46
 */
public class StreamSST extends SharedStringsTable implements AutoCloseable {
    private final FileBackedList<StreamCTRst> list;

    private StreamSST(PackagePart part, File file, int cacheSize) throws IOException {
        this.list = new FileBackedList<>(StreamCTRst.class, file, cacheSize);
        readFrom(part.getInputStream());
    }

    public static StreamSST getSharedStringTable(File file, int sstCacheSize, OPCPackage opCpkg) throws IOException {
        List<PackagePart> parts = opCpkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
        return parts.size() == 0 ? null : new StreamSST(parts.get(0), file, sstCacheSize);
    }

    /**
     * @Description: read the file named "sharedString.xml".
     *                 parse this file using XMLEventReader
     * @params [is]
     * @return void
     * @throws IOException
     */
    @Override
    public void readFrom(InputStream is) throws IOException {
        try {
            XMLEventReader xmlEventReader = StaxHelper.newXMLInputFactory().createXMLEventReader(is);

            while(xmlEventReader.hasNext()) {
                XMLEvent xmlEvent = xmlEventReader.nextEvent();

                if(xmlEvent.isStartElement() && xmlEvent.asStartElement().getName().getLocalPart().equals("si")) {
                    list.add(parseCT_Rst(xmlEventReader));
                }
            }
        } catch(XMLStreamException e) {
            throw new IOException(e);
        }
    }

    /**
     * Parses a {@code <si>} String Item. Returns just the text and drops the formatting.
     * See <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sharedstringitem.aspx">xmlschema type {@code CT_Rst}</a>.
     */
    private StreamCTRst parseCT_Rst(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <si>;  Post condition: pointing to </si>
        StringBuilder buf = new StringBuilder();
        XMLEvent xmlEvent;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch(xmlEvent.asStartElement().getName().getLocalPart()) {
                case "t": // Text
                    buf.append(xmlEventReader.getElementText());
                    break;
                case "r": // Rich Text Run
                    parseCT_RElt(xmlEventReader, buf);
                    break;
                case "rPh": // Phonetic Run
                case "phoneticPr": // Phonetic Properties
                    skipElement(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
        return buf.length() > 0 ? new StreamCTRst(buf.toString()) : null;
    }

    /**
     * Parses a {@code <r>} Rich Text Run. Returns just the text and drops the formatting.
     * See <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.run.aspx">xmlschema type {@code CT_RElt}</a>.
     */
    private void parseCT_RElt(XMLEventReader xmlEventReader, StringBuilder buf) throws XMLStreamException {
        // Precondition: pointing to <r>;  Post condition: pointing to </r>
        XMLEvent xmlEvent;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch(xmlEvent.asStartElement().getName().getLocalPart()) {
                case "t": // Text
                    buf.append(xmlEventReader.getElementText());
                    break;
                case "rPr": // Run Properties
                    skipElement(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
    }

    /**
     * for startElement, skip it.
     * @param xmlEventReader
     * @throws XMLStreamException
     */
    private void skipElement(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to start element;  Post condition: pointing to end element
        while(xmlEventReader.nextTag().isStartElement()) {
            skipElement(xmlEventReader); // recursively skip over child
        }
    }

    public CTRst getEntryAt(int idx) {
        CTRst result = list.getAt(idx);
        return result != null ? result : StreamCTRst.EMPTY;
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
        list.close();
    }

}
