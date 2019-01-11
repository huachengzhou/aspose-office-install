package com.help;


import java.io.File;

/**
 * @Auther: zch
 * @Date: 2019/1/11 16:09
 * @Description:当前路径处理
 */
public class TestFile {

    public static String getTestDataDir(Class<?> c) {
        File dir = new File(System.getProperty("user.dir"));
        StringBuilder builder = new StringBuilder(128);
        builder.append(dir.toString()).append(File.separator).append("src").append("\\test\\java\\");
        int i = 0;
        for (String s : c.getName().split("\\.")) {
            i++;
            if (i == c.getName().split("\\.").length){
                builder.append(s.toLowerCase()).append("\\");
            }else {
                builder.append(s).append("\\");
            }
        }
        dir = new File(builder.toString());
        if (!dir.isDirectory()) {
            dir.mkdir();
        }
        System.out.println("files:"+builder.toString());
        return builder.toString();
    }

}
