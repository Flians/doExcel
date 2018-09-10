package com.unaware.poi.excel.streamreader;

import com.unaware.poi.excel.exception.ParseException;
import com.unaware.poi.excel.ssimpl.StreamCell;
import com.unaware.poi.excel.ssimpl.StreamRow;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.events.*;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author Unaware
 * @Title: SheetReader
 * @ProjectName excel
 * @Description: Parse the xml node using XMLEventReader line by line
 *                 There are some suggestions:
 *                 Using "XMLStreamReader" to parse documents is about 30% faster than "XMLEventReader".
 *                 So maybe you can try it if you are free
 * @date 2018/7/12 15:23
 */
public class SheetReader implements Iterable<Row>, AutoCloseable {
    private final SharedStringsTable sharedStringSource;
    private final StylesTable stylesSource;
    private final XMLEventReader parser;
    private final XMLStreamReader mergedReader;
    private final List<CellRangeAddress> mergedRegions = new ArrayList<>();
    private final List<Double> colWidth = new ArrayList<>();


    private List<Row> rowCache = new ArrayList<>();
    private Iterator<Row> rowCacheIterator;

    private int rowCacheSize;
    private int numMergedRegions = 0;
    private int firstRowNum;
    private int lastRowNum;
    private int currentRowNum;
    private int firstColNum = 0;
    private int currentColNum;
    private boolean use1904Dates;
    private String lastContents;
    private StreamRow currentRow;
    private StreamCell currentCell;

    /**
     * constructor
     * @param sharedStringSource
     * @param stylesSource
     * @param inputStream
     * @param mergedStream
     * @param use1904Dates
     * @param rowCacheSize
     * @throws XMLStreamException
     */
    SheetReader(SharedStringsTable sharedStringSource, StylesTable stylesSource, InputStream inputStream, InputStream mergedStream, boolean use1904Dates, int rowCacheSize) throws XMLStreamException {
        this.sharedStringSource = sharedStringSource;
        this.stylesSource = stylesSource;
        this.parser = StaxHelper.newXMLInputFactory().createXMLEventReader(inputStream);
        this.mergedReader = StaxHelper.newXMLInputFactory().createXMLStreamReader(mergedStream);
        this.rowCacheSize = rowCacheSize;
        this.use1904Dates = use1904Dates;
        // obtain the mergeCells
        // we can't get merged cells information until sheet is parsed.
        // So we traversed sheet in advance to get merged cells information.
        parseMergedCells();
    }

    /**
     * obtain the mergeCells using XMLStreamReader
     * @throws XMLStreamException
     */
    private void parseMergedCells() throws XMLStreamException {
        int i;
        // Loop over XML input stream and process events
        while (mergedReader.hasNext()) {
            if (mergedReader.isStartElement()) {
                // obtain the mergeCell
                if ("mergeCell".equalsIgnoreCase(mergedReader.getLocalName())) {
                    for (i = 0; i < mergedReader.getAttributeCount(); i++) {
                        if ("ref".equals(mergedReader.getAttributeName(i).getLocalPart())) {
                            mergedRegions.add(CellRangeAddress.valueOf(mergedReader.getAttributeValue(i)));
                            break;
                        }
                    }
                    // obtain the number of mergeCells
                } else if ("mergeCells".equalsIgnoreCase(mergedReader.getLocalName())) {
                    for (i = 0; i < mergedReader.getAttributeCount(); i++) {
                        if ("count".equals(mergedReader.getAttributeName(i).getLocalPart())) {
                            numMergedRegions = Integer.parseInt(mergedReader.getAttributeValue(i));
                            break;
                        }
                    }
                }
            }
            mergedReader.next();
        }
        mergedReader.close();
    }

    /**
     * @Description: read through a number of rows equal to the rowCacheSize
     *                 Or until there is no more data to read
     * @return boolean
     */
    private boolean getRow() {
        try {
            rowCache.clear();
            while (rowCache.size() < rowCacheSize && parser.hasNext()) {
                handleEvent(parser.nextEvent());
            }
            rowCacheIterator = rowCache.iterator();
            return rowCacheIterator.hasNext();
        } catch (XMLStreamException e) {
            throw new ParseException("Error reading XML stream", e);
        }

    }

    /**
     * parse the node to get the cell information
     * @param xmlEvent
     */
    private void handleEvent(XMLEvent xmlEvent) {
        if (xmlEvent.getEventType() == XMLStreamConstants.CHARACTERS) {
            Characters c = xmlEvent.asCharacters();
            lastContents += c.getData();
        } else if (xmlEvent.getEventType() == XMLStreamConstants.START_ELEMENT &&
                isSpreadsheetTag(xmlEvent.asStartElement().getName())) {
            StartElement startElement = xmlEvent.asStartElement();
            String tagLocalName = startElement.getName().getLocalPart();

            switch (tagLocalName) {
                case "row":
                    Attribute rowNumAttr = startElement.getAttributeByName(new QName("r"));
                    int rowIndex = currentRowNum;
                    if (rowNumAttr != null) {
                        rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
                        currentRowNum = rowIndex;
                    }
                    currentRow = new StreamRow(rowIndex);
                    currentColNum = firstColNum;
                    break;
                case "col":
                    Attribute widthAttr = startElement.getAttributeByName(new QName("width"));
                    if (widthAttr != null) {
                        colWidth.add(Double.parseDouble(widthAttr.getValue()));
                    }
                    break;
                case "c":
                    //obtain the index of cell
                    Attribute rAttr = startElement.getAttributeByName(new QName("r"));
                    if (rAttr != null) {
                        CellRangeAddress temp = CellRangeAddress.valueOf(rAttr.getValue());
                        currentCell = new StreamCell(temp.getFirstColumn(), temp.getFirstRow(), use1904Dates);
                    } else {
                        currentCell = new StreamCell(currentColNum, currentRowNum, use1904Dates);
                    }
                    setFormatString(startElement, currentCell);

                    //obtain the type of the cell data
                    Attribute type = startElement.getAttributeByName(new QName("t"));
                    if (type != null) {
                        currentCell.setType(type.getValue());
                    } else {
                        currentCell.setType("n");
                    }

                    //obtain the style of the cell
                    Attribute style = startElement.getAttributeByName(new QName("s"));
                    if (style != null) {
                        String indexStr = style.getValue();
                        try {
                            int index = Integer.parseInt(indexStr);
                            currentCell.setCellStyle(stylesSource.getStyleAt(index));
                        } catch (NumberFormatException e) {
                            System.out.println("Warn: Ignoring invalid style index {}" + indexStr);
                        }
                    } else {
                        currentCell.setCellStyle(stylesSource.getStyleAt(0));
                    }
                    break;
                case "dimension":
                    Attribute refAttr = startElement.getAttributeByName(new QName("ref"));
                    String ref = refAttr != null ? refAttr.getValue() : null;
                    if (ref != null) {
                        // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
                        CellRangeAddress dTemp = CellRangeAddress.valueOf(ref);
                        firstColNum = dTemp.getFirstColumn();
                        firstRowNum = dTemp.getFirstRow();
                        lastRowNum = dTemp.getLastRow();
                    }
                    break;
                case "f":
                    if (currentCell != null) {
                        currentCell.setType("str");
                    }
                    break;
            }
            // Clear contents cache
            lastContents = "";
        } else if (xmlEvent.getEventType() == XMLStreamConstants.END_ELEMENT
                && isSpreadsheetTag(xmlEvent.asEndElement().getName())) {
            EndElement endElement = xmlEvent.asEndElement();
            String tagLocalName = endElement.getName().getLocalPart();

            switch (tagLocalName) {
                case "v":
                case "t":
                    currentCell.setRawContents(unformattedContents());
                    break;
                case "row":
                    if (currentRow != null) {
                        rowCache.add(currentRow);
                        currentRowNum++;
                    }
                    break;
                case "c":
                    currentRow.getCellMap().put(currentCell.getColumnIndex(), currentCell);
                    currentCell = null;
                    currentColNum++;
                    break;
                case "f":
                    if (currentCell != null) {
                        currentCell.setFormula(lastContents);
                    }
                    break;
            }
        }
    }

    /**
     * Read the numeric format string out of the styles table for this cell. Stores
     * the result in the Cell.
     *
     * @param startElement
     * @param cell
     */
    private void setFormatString(StartElement startElement, StreamCell cell) {
        Attribute cellStyle = startElement.getAttributeByName(new QName("s"));
        String cellStyleString = (cellStyle != null) ? cellStyle.getValue() : null;
        XSSFCellStyle style = null;

        if (cellStyleString != null) {
            style = stylesSource.getStyleAt(Integer.parseInt(cellStyleString));
        } else if (stylesSource.getNumCellStyles() > 0) {
            style = stylesSource.getStyleAt(0);
        }

        if (style != null) {
            cell.setNumericFormatIndex(style.getDataFormat());
            String formatString = style.getDataFormatString();

            if (formatString != null) {
                cell.setNumericFormat(formatString);
            } else {
                cell.setNumericFormat(BuiltinFormats.getBuiltinFormat(cell.getNumericFormatIndex()));
            }
        } else {
            cell.setNumericFormatIndex(null);
            cell.setNumericFormat(null);
        }
    }

    /**
     * Returns true if a tag is part of the main namespace for SpreadsheetML:
     * <ul>
     * <li>http://schemas.openxmlformats.org/spreadsheetml/2006/main
     * <li>http://purl.oclc.org/ooxml/spreadsheetml/main
     * </ul>
     * As opposed to http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing, etc.
     *
     * @param name
     * @return
     */
    private boolean isSpreadsheetTag(QName name) {
        return (name.getNamespaceURI() != null
                && name.getNamespaceURI().endsWith("/main"));
    }

    /**
     * Returns the contents of the cell, with no formatting applied
     *
     * @return
     */
    private String unformattedContents() {
        switch (currentCell.getType()) {
            case "s":           //string stored in shared table
                if (!lastContents.isEmpty()) {
                    int idx = Integer.parseInt(lastContents);
                    return new XSSFRichTextString(sharedStringSource.getEntryAt(idx)).toString();
                }
                return lastContents;
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(lastContents).toString();
            default:
                return lastContents;
        }
    }

    /**
     * return all the merged regions of this sheet
     *
     * @return
     */
    public List<CellRangeAddress> getMergedRegions() {
        return mergedRegions;
    }

    /**
     * return the number of all merged regions
     *
     * @return
     */
    public int getNumMergedRegions() {
        return numMergedRegions;
    }

    /**
     * return the index of the first row which is not empty
     *
     * @return
     */
    public int getFirstRowNum() {
        return firstRowNum;
    }

    /**
     * return the index of the last row which is not empty
     *
     * @return
     */
    public int getLastRowNum() {
        return lastRowNum;
    }

    /**
     * Return the width of each column which is from worksheet/cols/col in sheet_.xml
     *
     * @return
     */
    public List<Double> getColWidth() {
        return colWidth;
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
    public void close() throws Exception {
        parser.close();
        mergedReader.close();
        mergedRegions.clear();
        colWidth.clear();
        rowCache.clear();
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Row> iterator() {
        return new StreamRowIterator();
    }

    private class StreamRowIterator implements Iterator<Row> {
        StreamRowIterator() {
            if (rowCacheIterator == null) {
                hasNext();
            }
        }

        @Override
        public boolean hasNext() {
            return (rowCacheIterator != null && rowCacheIterator.hasNext()) || getRow();
        }

        @Override
        public Row next() {
            return rowCacheIterator.next();
        }

        @Override
        public void remove() {
            throw new RuntimeException("NotSupported");
        }
    }

    /**
     * obtain the mergeCells from InputStream using XMLReader
     * @param mergedStream
     */
    private void parseMergedCellsByXMLReader (InputStream mergedStream) {
        try {
            XMLReader xmlReader = XMLReaderFactory.createXMLReader();
            ObtainMergedRegions obtainMergedRegions = new ObtainMergedRegions();
            xmlReader.setContentHandler(obtainMergedRegions);
            xmlReader.parse(new InputSource(mergedStream));
            mergedStream.close();
        } catch (SAXException | IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * obtain the mergeCells using XMLReader
     */
    private class ObtainMergedRegions extends DefaultHandler {
        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            String value = attributes.getValue("ref");
            if ("mergeCell".equalsIgnoreCase(name) && value != null) {
                mergedRegions.add(CellRangeAddress.valueOf(value));
            } else if ("mergeCells".equalsIgnoreCase(name) && value != null) {
                numMergedRegions = Integer.parseInt(value);
            }
        }
    }
}