package org.dovnard.visio;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * User: DovnarDmitriy
 * Date: 11.11.19
 * Time: 15:57
 */
public class Connect implements XmlContent {
    private String fromSheet;
    private String fromCell;
    private String fromPart;
    private String toSheet;
    private String toCell;
    private String toPart;

    public String getFromSheet() {
        return fromSheet;
    }

    public void setFromSheet(String fromSheet) {
        this.fromSheet = fromSheet;
    }

    public String getFromCell() {
        return fromCell;
    }

    public void setFromCell(String fromCell) {
        this.fromCell = fromCell;
    }

    public String getFromPart() {
        return fromPart;
    }

    public void setFromPart(String fromPart) {
        this.fromPart = fromPart;
    }

    public String getToSheet() {
        return toSheet;
    }

    public void setToSheet(String toSheet) {
        this.toSheet = toSheet;
    }

    public String getToCell() {
        return toCell;
    }

    public void setToCell(String toCell) {
        this.toCell = toCell;
    }

    public String getToPart() {
        return toPart;
    }

    public void setToPart(String toPart) {
        this.toPart = toPart;
    }

    @Override
    public Element getXmlElement(Document document) {
        Element root = document.createElement("Connect");
        root.setAttribute("FromSheet", fromSheet);
        root.setAttribute("FromCell", fromCell);
        root.setAttribute("FromPart", fromPart);
        root.setAttribute("ToSheet", toSheet);
        root.setAttribute("ToCell", toCell);
        root.setAttribute("ToPart", toPart);
        return root;
    }
}
