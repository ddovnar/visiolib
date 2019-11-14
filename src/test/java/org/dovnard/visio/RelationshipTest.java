package org.dovnard.visio;

import junit.framework.Assert;
import org.junit.Test;
import org.w3c.dom.Document;

import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 07.11.19
 * Time: 0:53
 */
public class RelationshipTest {
    @Test
    public void test() {
        Relationship rels = new Relationship();
        try {
            rels.addRelation("id1", "test1", "todo1");
            rels.addRelation("id2", "test2", "todo2");

            Document xml = rels.toXml();
            printXml(xml);
        } catch (Exception e) {
            e.printStackTrace();
            Assert.fail();
        }
    }

    public void printXml(Document document) throws TransformerException {
        DOMSource source = new DOMSource(document);
        StreamResult result = new StreamResult(System.out);

        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
        transformer.transform(source, result);
    }
}
