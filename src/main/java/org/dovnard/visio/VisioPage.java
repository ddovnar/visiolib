package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * Visio project
 * Author: Dmitriy Dovnar
 * Date: 07.11.19
 * Time: 0:14
 */
public interface VisioPage extends XmlExportable {
    String getName();
    //void setName(String name);
    Element getXmlElement(Document document);
    void setId(int id);
    void setRelId(int relId);

    int getId();
    String getRelId();
    String getShortFileName();

    Shape addShape(ShapeType shapeType, String name, double xpos, double ypos);
    Shape addShape(ShapeType shapeType, String name, String title, double xpos, double ypos);

    void addConnectorBottomTop(Shape shapeSource, Shape shapeTarget);
}
