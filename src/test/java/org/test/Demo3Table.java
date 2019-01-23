package org.test;

import com.aspose.words.*;
import com.google.common.collect.Sets;
import com.help.TestFile;
import com.red.tool.other.MergeCellModel;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.UUID;

/**
 * @Auther: zch
 * @Date: 2019/1/23 15:11
 * @Description:
 */
public class Demo3Table {

    @Test
    public void testA() throws Exception {
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString().substring(0,4));
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("估价结果一览表");
        int i = 0;
        while (i < stringList.size() + 1) {
            Set<MergeCellModel> mergeCellModelList = Sets.newHashSet();
            Table table = builder.startTable();
            switch (i) {
                case 1:
                    for (int k = 0; k < 9; k++) {
                        switch (k) {
                            case 0:
                                builder.insertCell();
                                builder.writeln("产权证号");
                                break;
                            case 1:
                                builder.insertCell();
                                builder.writeln("业务件号");
                                break;
                            case 2:
                                builder.insertCell();
                                builder.writeln("房屋座落");
                                break;
                            case 3:
                                builder.insertCell();
                                builder.writeln("权利人");
                                break;
                            case 4:
                                builder.insertCell();
                                builder.writeln("共有方式");
                                break;
                            case 5:
                                builder.insertCell();
                                builder.writeln("规划用途");
                                break;
                            case 6:
                                builder.insertCell();
                                builder.writeln("所在层数");
                                break;
                            case 7:
                                builder.insertCell();
                                builder.writeln("建筑面积");
                                break;
                            case 8:
                                builder.insertCell();
                                builder.writeln("房地产单价");
                                break;
                            case 9:
                                builder.insertCell();
                                builder.writeln("房地产价值");
                                break;
                            default:
                                builder.insertCell();
                                break;
                        }
                    }
                    builder.endRow();
                    break;
                case 2:
                    this.testA2(builder, stringList.get(i - 1));
                    break;
                case 3:
                    this.testA2(builder, stringList.get(i - 1));
                    break;
                default:
                    break;
            }

            i++;

            builder.getCellFormat().getBorders().getLeft().setLineWidth(1.0);
            builder.getCellFormat().getBorders().getRight().setLineWidth(1.0);
            builder.getCellFormat().getBorders().getTop().setLineWidth(1.0);
            builder.getCellFormat().getBorders().getBottom().setLineWidth(1.0);
            builder.getCellFormat().setWidth(100);
            builder.getCellFormat().setVerticalMerge(CellVerticalAlignment.CENTER);
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
            builder.endTable();
        }

        //当前类 和 当前运行方法
        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), Thread.currentThread().getStackTrace()[1].getMethodName(), ".docx"));
    }


    private void testA2(DocumentBuilder builder, String str) throws Exception {
        for (int k = 0; k < 9; k++) {
            switch (k) {
                case 0:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 1:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 2:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 3:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 4:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 5:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 6:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 7:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 8:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                case 9:
                    builder.insertCell();
                    builder.writeln(str);
                    break;
                default:
                    builder.insertCell();
                    break;
            }
        }
        builder.endRow();
    }

}
