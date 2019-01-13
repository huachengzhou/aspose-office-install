package com.featurescomparison.workingwithranges.insertbeforeandafterranges;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Author noatn
 * @Description
 * @createDate 2019/1/13
 **/
public class AsposeInsert {

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "data\\";
        Document doc = new Document(dataPath + "document.doc");

        // This gets a live collection of all shape nodes in the document.
        NodeCollection shapeCollection = doc.getChildNodes(NodeType.SHAPE, true);

        // Since we will be adding/removing nodes, it is better to copy all collection
        // into a fixed size array, otherwise iterator will be invalidated.
        Node[] shapes = shapeCollection.toArray();

        for (Node node : shapes)
        {
            Shape shape = (Shape)node;
            // Filter out all shapes that we don't need.
            if (shape.getShapeType() == ShapeType.TEXT_BOX)
            {
                // Create a new shape that will replace the existing shape.
                Shape image = new Shape(doc, ShapeType.IMAGE);

                // Load the image into the new shape.
                image.getImageData().setImage(dataPath + "background.jpg");

                // Make new shape's position to match the old shape.
                image.setLeft(shape.getLeft());
                image.setTop(shape.getTop());
                image.setWidth(shape.getWidth());
                image.setHeight(shape.getHeight());
                image.setRelativeHorizontalPosition(shape.getRelativeHorizontalPosition());
                image.setRelativeVerticalPosition(shape.getRelativeVerticalPosition());
                image.setHorizontalAlignment(shape.getHorizontalAlignment());
                image.setVerticalAlignment(shape.getVerticalAlignment());
                image.setWrapType(shape.getWrapType());
                image.setWrapSide(shape.getWrapSide());

                // Insert new shape after the old shape and remove the old shape.
                shape.getParentNode().insertAfter(image, shape);
                shape.remove();
            }
        }

        doc.save(dataPath + "AsposeReplaceTextboxesWithImages_Out.doc");
        System.out.println("Process Completed Successfully");
    }

}
