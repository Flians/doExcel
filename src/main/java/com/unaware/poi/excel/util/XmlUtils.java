package com.unaware.poi.excel.util;

import com.unaware.poi.excel.exception.ParseException;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Unaware
 * @Title: XmlUtils
 * @ProjectName excel
 * @Description: parse XML to generate {@link Document}, support for lookups
 * @date 2018/7/12 15:21
 */
public class XmlUtils {

    /**
     * @Description: parse XML from inputStream to generate {@link Document}
     * @params is
     * @return org.w3c.dom.Document
     * @throws
     */
    public static Document document(InputStream is) {
        try {
            return DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(is);
        } catch (SAXException | IOException | ParserConfigurationException e) {
            throw new ParseException(e);
        }
    }

    /**
     * @Description: look up Nodes with the parameter xpath from {@link Document}.
     *               The parameter xpath represents the path of node.
     * @params [document, xpath]
     * @return org.w3c.dom.NodeList
     * @throws
     */
    public static NodeList searchForNodeList(Document document, String xpath) {
        try {
            return (NodeList) XPathFactory.newInstance().newXPath().compile(xpath)
                    .evaluate(document, XPathConstants.NODESET);
        } catch (XPathExpressionException e) {
            throw new ParseException(e);
        }
    }

}
