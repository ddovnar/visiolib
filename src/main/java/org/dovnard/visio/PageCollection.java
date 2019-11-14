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
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * User: DovnarDmitriy
 * Date: 08.11.19
 * Time: 11:00
 */
public class PageCollection implements XmlExportable {
    private List<VisioPage> pageList;
    private Relationship pageRels;

    public PageCollection() {
        pageList = new ArrayList<>();
        pageRels = new Relationship();
    }

    public VisioPage createPage(String pageName, ContentTypes contentTypes) throws Exception {
        VisioPage newPage = new VisioPageImpl(pageName);
        newPage.setId(pageList.size() + 1);
        newPage.setRelId(pageList.size() + 1);
        pageList.add(newPage);

        pageRels.addRelation(newPage.getRelId(), newPage.getShortFileName(), "http://schemas.microsoft.com/visio/2010/relationships/page");

        contentTypes.registerPage(newPage.getFileName());

        return newPage;
    }

    public List<VisioPage> getPageList() {
        return pageList;
    }

    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("Pages");
        root.setAttribute("xmlns", "http://schemas.microsoft.com/office/visio/2012/main");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xml:space", "preserve");
        document.appendChild(root);

        Iterator<VisioPage> pageIterator = pageList.iterator();
        while (pageIterator.hasNext()) {
            VisioPage page = pageIterator.next();
            root.appendChild(page.getXmlElement(document));
        }
        return document;
    }

    @Override
    public String getFileName() {
        return "/visio/pages/" + "pages.xml";
    }

    private void writeAllPages(String path) throws ParserConfigurationException, TransformerException, IOException {
        Iterator<VisioPage> pageIterator = pageList.iterator();
        while (pageIterator.hasNext()) {
            VisioPage page = pageIterator.next();
            page.writeToFile(path);
        }

        pageRels.writeToFile(path + "/visio/pages/_rels/", "pages.xml.rels");
    }
    @Override
    public void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, getFileName(), toXml());
        writeAllPages(path);
    }

    @Override
    public void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, filename, toXml());
        writeAllPages(path);
    }
}
