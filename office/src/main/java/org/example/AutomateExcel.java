package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author cmsxyz@163.com
 * @date 2024-07-11 17:52
 * @usage 解决每周跬步日期编排问题
 */
public class AutomateExcel {

    public static void main(String[] args) throws IOException {

            // 指定Excel文件路径
            String filePath = "路径";

            // 创建文件输入流
            FileInputStream fis = new FileInputStream(new File(filePath));

            // 创建工作簿对象，根据文件类型选择正确的构造函数
            Workbook workbook = WorkbookFactory.create(fis);

            // 获取第五个工作表
            Sheet sheet = workbook.getSheetAt(4);

            // 获取第二行（行索引从0开始）
            Row row = sheet.getRow(1);

            // 获取第三列（列索引从0开始）
            Cell cell = row.getCell(2);

            // 根据单元格类型获取其值
            if (cell != null) {
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println("Cell value: " + cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.println("Cell value: " + cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.println("Cell value: " + cell.getBooleanCellValue());
                        break;
                    default:
                        System.out.println("Unsupported cell type.");
                }
            } else {
                System.out.println("Cell is null.");
            }

            // 关闭资源
            fis.close();
            workbook.close();

        }
    }

