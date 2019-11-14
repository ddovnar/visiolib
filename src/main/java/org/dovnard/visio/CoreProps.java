package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * User: DovnarDmitriy
 * Date: 07.11.19
 * Time: 15:24
 */
public class CoreProps implements XmlExportable {
    private LocalDateTime created;
    private LocalDateTime modified;
    private String language;
    private String title;
    private String subject;
    private String description;

    public CoreProps() {
        created = LocalDateTime.now();
        modified = LocalDateTime.now();
        language = "en-US";
        title = "";
        subject = "";
        description = "";
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public LocalDateTime getCreated() {
        return created;
    }

    public void setCreated(LocalDateTime created) {
        this.created = created;
    }

    public LocalDateTime getModified() {
        return modified;
    }

    public void setModified(LocalDateTime modified) {
        this.modified = modified;
    }

    public String getLanguage() {
        return language;
    }

    public void setLanguage(String language) {
        this.language = language;
    }

    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("cp:coreProperties");
        root.setAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        root.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
        root.setAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
        root.setAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
        root.setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
        document.appendChild(root);

        Element titleElement = document.createElement("dc:title");
        titleElement.setTextContent(title);
        root.appendChild(titleElement);

        Element subjectElement = document.createElement("dc:subject");
        subjectElement.setTextContent(subject);
        root.appendChild(subjectElement);

        root.appendChild(document.createElement("dc:creator"));
        root.appendChild(document.createElement("cp:keywords"));

        Element descElement = document.createElement("dc:description");
        descElement.setTextContent(description);
        root.appendChild(descElement);

        root.appendChild(document.createElement("cp:lastModifiedBy"));
        root.appendChild(document.createElement("cp:lastPrinted"));

        Element created = document.createElement("dcterms:created");
        created.setAttribute("xsi:type", "dcterms:W3CDTF");
        created.setTextContent(this.created.toString());
//        created.setTextContent("2016-08-29T06:42:32Z");
        root.appendChild(created);

        Element modified = document.createElement("dcterms:modified");
        modified.setAttribute("xsi:type", "dcterms:W3CDTF");
//        modified.setTextContent("2016-08-29T06:42:32Z");
        modified.setTextContent(this.modified.toString());
        root.appendChild(modified);

        root.appendChild(document.createElement("cp:category"));

        Element lang = document.createElement("dc:language");
//        lang.setTextContent("en-US");
        lang.setTextContent(language);
        root.appendChild(lang);

        return document;
    }

    @Override
    public String getFileName() {
        return "/docProps/" + "core.xml";
    }

    @Override
    public void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, getFileName(), toXml());
    }

    @Override
    public void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, filename, toXml());
    }
}
