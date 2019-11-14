# visiolib
Java library for create visio diagrams.

## Info
This project is in proggress, but currently have a lot of worked based functionality for creating simple diagrams. See example.

- draw rectangle shapes
- connect shapes. In top-to-bottom direction

# How to use

Example:
```java
...
VisioDocument visioDocument = null;
try {
    visioDocument = VisioDocumentBuilder.newDocument();
    VisioPage page = visioDocument.createPage("Page1");
    Shape rec1 = page.addShape(ShapeType.RECTANGLE, "test1", "root", 4.0, 10.0);
    Shape rec2 = page.addShape(ShapeType.RECTANGLE, "test2", "Child1", 1.0, 8.0);
    page.addConnectorBottomTop(rec1, rec2);
    
    visioDocument.createVisioFile(getUniqFileName("doc_"));
} catch (Exception e) {
    ...
}
...
```

Result:
```sh
                    +------+
                    | root |
                    +------+
                        |
        +---------------+
        |
    +------+
    |child1|
    +------+
```