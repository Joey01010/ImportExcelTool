package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static final int CELL_ROW = 376;
    public static final int CELL_IDX = 368;
    public static final String ROOT = System.getProperty("user.dir");
    public static final String PATH1 = "/模板/主文件录入模板.xlsx";
    public static final String PATH2 = "/工艺卡/工艺卡.xlsx";
    public static final String PATH3 = "/数据/半成品主文件导入数据.xlsx";

    /**
     * 获取模板数据方法
     *
     * @return
     * @throws IOException
     */
    public static List<List<String>> getTemplate() throws IOException {

        //读取模板表格数据
        File template = new File(ROOT + PATH1);
        Workbook wb = WorkbookFactory.create(template);
        Sheet tSheet = wb.getSheet("物料#物料(FBillHead)");

        //每条记录的模板参数
        List<String> title1 = new ArrayList<>();
        List<String> title2 = new ArrayList<>();
        List<String> param1 = new ArrayList<>();
        List<String> param2 = new ArrayList<>();
        List<String> param3 = new ArrayList<>();
        List<String> param4 = new ArrayList<>();
        List<String> param5 = new ArrayList<>();

        //遍历实体数据
        List<List<String>> lists = Arrays.asList(title1, title2, param1);
        for (int i = 0; i < lists.size(); i++) {
            Row row = tSheet.getRow(i);
            for (int j = 0; j < CELL_ROW; j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    lists.get(i).add("");
                } else {
                    cell.setCellType(CellType.STRING);
                    lists.get(i).add(cell.getStringCellValue());
                }
            }
        }

        int i = 2;

        //循环二到五行
        List<List<String>> param = Arrays.asList(param2, param3, param4, param5);
        for (List<String> pm : param) {
            i = i + 1;
            Row row = tSheet.getRow(i);
            for (int k = CELL_IDX; k < CELL_ROW; k++) {
                Cell cell = row.getCell(k);
                if (cell != null) {
                    cell.setCellType(CellType.STRING);
                    String value = cell.getStringCellValue();
                    pm.add(value);
                }
            }
        }

        List<List<String>> result = Arrays.asList(title1, title2, param1, param2, param3, param4, param5);
        wb.close();
        return result;
    }

    /**
     * 获取工艺卡数据方法
     *
     * @return
     * @throws IOException
     */
    public static List<String> getData() throws IOException {
        List data = new ArrayList();
        //去重数组
        List temp = new ArrayList();
        File file = new File(ROOT + PATH2);
        Workbook wb = WorkbookFactory.create(file);
        //添加项目名称
        data.add(wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            String pdName = sheet.getRow(0).getCell(0).getStringCellValue();
            for (int j = 3; j < sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if(row == null){
                    continue;
                }
                Cell cell = row.getCell(1);
                if (cell == null || cell.getStringCellValue().equals("") || temp.contains(cell.getStringCellValue())) {
                    continue;
                }
                temp.add(cell.getStringCellValue());
                data.add(pdName + "-" + cell.getStringCellValue());
            }
        }
        wb.close();
        return data;
    }


    /**
     * 程序入口
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {
        int idx = 102440;

        List<List<String>> template = getTemplate();
        List<String> data = getData();

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("物料#物料(FBillHead)");

        //打印标题行
        for (int i = 0; i < 2; i++) {
            Row row = sheet.createRow(i);
            List<String> title = template.get(i);
            for (int j = 0; j < title.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(title.get(j));
            }
        }

        //打印每个产品的数据
        String pjName = data.get(0);
        for (int i = 1; i < data.size(); i++) {
            List<String> param1 = template.get(2);
            //遍历模板数据
            Row row = sheet.createRow((i - 1) * 5 + 2);
            for (int j = 0; j < param1.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(param1.get(j));

                String value = data.get(i);
                int strIdx = value.lastIndexOf("-");
                String pdName = value.substring(0, strIdx);
                String code = value.substring(strIdx + 1);
                String group = code.substring(0, 2);

                //录入规则
                if (j == 0) {
                    idx = idx + 1;
                    cell.setCellValue(idx);
                } else if (j == 5) {
                    cell.setCellValue(code);
                } else if (j == 6 || j == 7) {
                    cell.setCellValue(pdName + "切线压接看板");
                } else if (j == 9) {
                    if (group.equals("KC")) {
                        cell.setCellValue("KB" + code.substring(2));
                    } else {
                        cell.setCellValue(code);
                    }
                } else if (j == 11) {
                    cell.setCellValue(group);
                } else if (j == 12) {
                    switch (group) {
                        case "KB":
                            cell.setCellValue("一般过程，单次加工");
                            break;
                        case "KC":
                            cell.setCellValue("由KS合成的多合一看板");
                            break;
                        case "KD":
                            cell.setCellValue("由全自动插线设备加工的合成看板");
                            break;
                        case "KS":
                            cell.setCellValue("中间过程，多合一的单线或KB/KC下级半成品");
                            break;
                    }
                } else if (j == 18) {
                    cell.setCellValue(pjName);
                }
            }

            //打印2-5行数据
            for (int k = 0; k < 4; k++) {
                Row nextRow = sheet.createRow((i - 1) * 5 + k + 3);
                List<String> param = template.get(k + 3);
                for (int l = CELL_IDX; l < CELL_ROW; l++) {
                    Cell cell = nextRow.createCell(l);
                    cell.setCellValue(param.get(l - CELL_IDX));
                }
            }
        }

        FileOutputStream out = new FileOutputStream(ROOT + PATH3);
        wb.write(out);
        wb.close();
    }
}