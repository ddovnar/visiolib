package org.dovnard.visio;

/**
 * User: DovnarDmitriy
 * Date: 11.11.19
 * Time: 11:39
 */
public class ShapeFactory {
    public static Shape createShape(ShapeType shapeType, String name, double xpos, double ypos) {
        if (shapeType == ShapeType.RECTANGLE) {
            return new RectangleShape(name, xpos, ypos);
        }
        return null;
    }
}
