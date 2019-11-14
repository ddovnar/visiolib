package org.dovnard.visio;

import org.w3c.dom.Document;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 06.11.19
 * Time: 23:50
 */
public interface XmlExportable {
    Document toXml() throws ParserConfigurationException;
    String getFileName();
    void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException;
    void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException;
}
