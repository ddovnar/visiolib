package org.dovnard.visio;

/**
 * User: DovnarDmitriy
 * Date: 11.11.19
 * Time: 11:41
 */
public class RectangleShape extends Shape {
    public RectangleShape(String name, double xpos, double ypos) {
        super(name, xpos, ypos);
        setWidth(1.5);
        setHeight(1.0);
    }
}
