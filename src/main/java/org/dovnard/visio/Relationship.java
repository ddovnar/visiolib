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
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 07.11.19
 * Time: 0:31
 */
public class Relationship implements XmlExportable {
    private final String docSchema = "http://schemas.openxmlformats.org/package/2006/relationships";
    private Map<String, RelationshipItem> relations = new LinkedHashMap<>();

    public void addRelation(String id, String target, String type) throws Exception {
        if (relations.containsKey(id)) {
            throw new Exception("Relation with id='" + id + "' already exists");
        }
        RelationshipItem relItem = new RelationshipItem();

        relItem.setId(id);
        relItem.setTarget(target);
        relItem.setType(type);

        relations.put(id, relItem);
    }
    public void removeRelation(String id) {
        relations.remove(id);
    }

    @Override
    public Document toXml() throws ParserConfigurationException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();

        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document document = docBuilder.newDocument();
        document.setXmlStandalone(true);

        Element root = document.createElement("Relationships");
        root.setAttribute("xmlns", docSchema);
        document.appendChild(root);

        relations.forEach((id, item) -> {
            Element rel = document.createElement("Relationship");
            rel.setAttribute("Id", item.getId());
            rel.setAttribute("Target", item.getTarget());
            rel.setAttribute("Type", item.getType());

            root.appendChild(rel);
        });

        return document;
    }

    @Override
    public void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, getFileName(), toXml());
    }

    @Override
    public void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, filename, toXml());
    }

    @Override
    public String getFileName() {
        return "/_rels/" + ".rels";
    }

    final class RelationshipItem {
        private String id;
        private String target;
        private String type;

        String getId() {
            return id;
        }

        void setId(String id) {
            this.id = id;
        }

        String getTarget() {
            return target;
        }

        void setTarget(String target) {
            this.target = target;
        }

        String getType() {
            return type;
        }

        void setType(String type) {
            this.type = type;
        }
    }
}
