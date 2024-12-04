package com.customer.spring.util;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

import java.io.FileWriter;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.List;

/**
 * <p></p>
 *
 * @author hm 2024/12/3
 * @version V1.0
 * @modificationHistory=========================逻辑或功能性重大变更记录
 * @modify by user: hm 2024/12/3
 * @modify by reason:{方法名}:{原因}
 **/
public class FileUtil {

    public static void writeFile(String html, String parentPath, String htmlName) {
        FileWriter writer = null;
        try {
            String path = MessageFormat.format("{0}\\{1}", parentPath, htmlName);
            writer = new FileWriter(path);
            writer.write(html);
        } catch (Exception ignore) {
        } finally {
            if (writer != null) {
                try {
                    writer.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }

    public static boolean writeData(List<List<String>> rows, String name, List<String> header) {
        ExcelWriter writer = ExcelUtil.getWriter(MessageFormat.format("F:\\project\\data\\collect\\CollectData\\src\\main\\resources\\static\\file\\{0}", name));
        if (!name.endsWith(".xlsx")) {
            name = MessageFormat.format("{0}.xlsx", name);
        }
        writer.setSheet(name);
        writer.writeHeadRow(header);
        writer.write(rows);
        writer.setColumnWidth(-1, 36);
        writer.flush();
        writer.close();
        return true;
    }
}
