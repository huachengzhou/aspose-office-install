package com.featurescomparison.workingwithimages;

import com.help.TestFile;
import org.apache.poi.hwpf.HWPFDocument;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.hwpf.usermodel.Picture;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:57
 * @Description:
 */
public class ApacheExtractImages {

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "data\\";
        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                dataPath + "document.doc"));
        List<Picture> pics = doc.getPicturesTable().getAllPictures();

        for (int i = 0; i < pics.size(); i++) {
            Picture pic = (Picture) pics.get(i);

            FileOutputStream outputStream = new FileOutputStream(
                    dataPath + "Apache_"
                            + pic.suggestFullFileName());
            outputStream.write(pic.getContent());
            outputStream.close();
        }
    }

}
