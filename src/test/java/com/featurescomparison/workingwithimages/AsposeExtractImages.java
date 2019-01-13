package com.featurescomparison.workingwithimages;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:57
 * @Description:
 */
public class AsposeExtractImages {

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "data\\";

        Document doc = new Document(dataPath + "document.doc");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        int imageIndex = 0;
        for (Shape shape : (Iterable<Shape>) shapes) {
            if (shape.hasImage()) {
                String imageFileName = java.text.MessageFormat.format(
                        "Aspose.Images.{0}{1}", imageIndex, FileFormatUtil
                                .imageTypeToExtension(shape.getImageData()
                                        .getImageType()));
                shape.getImageData().save(dataPath + imageFileName);

                imageIndex++;
            }
        }
    }

}
