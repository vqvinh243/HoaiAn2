package com.mycompany.hoaian2;

import com.mycompany.models.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import javax.swing.table.TableModel;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Controller {
    private Processor processor = new Processor();

    public Processor getProcessor() {
        return processor;
    }

    public static HSSFCellStyle createStyleForSum(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)11);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    public void exportFile(Page page, String fileName, TableModel tableModel) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        Order order = new Order();
        HSSFSheet sheet = workbook.createSheet("don hang");
        sheet.setColumnWidth(Column.STT.getIndex(), 1500);
        sheet.setColumnWidth(Column.TEN_SAN_PHAM.getIndex(), 4500);
        sheet.setColumnWidth(Column.DON_CHINH.getIndex(), 3500);
        sheet.setColumnWidth(Column.L1.getIndex(), 3500);
        sheet.setColumnWidth(Column.L2.getIndex(), 3500);
        sheet.setColumnWidth(Column.TONG.getIndex(), 3500);
        sheet.setColumnWidth(Column.DON_GIA.getIndex(), 3500);
        sheet.setColumnWidth(Column.THANH_TIEN.getIndex(), 3500);
        sheet.setColumnWidth(Column.LOI_NHUAN.getIndex(), 3500);

        int rownum = 0;
        Cell cell;
        Row row;
        //
        HSSFCellStyle titleStyle = Controller.createStyleForTitle(workbook);
        HSSFCellStyle normalCellStyle = Controller.createStyleForNormalCell(workbook);
        HSSFCellStyle productNameCellStyle = Controller.createStyleForProductName(workbook);
        HSSFCellStyle sumCellStyle = Controller.createStyleForSum(workbook);

        row = sheet.createRow(rownum);

        // STT
        cell = row.createCell(Column.STT.getIndex(), CellType.STRING);
        cell.setCellValue("");
        // Tên sản phẩm
        cell = row.createCell(Column.TEN_SAN_PHAM.getIndex(), CellType.STRING);
        cell.setCellValue(Column.TEN_SAN_PHAM.getName());
        cell.setCellStyle(titleStyle);
        // Đơn chính
        cell = row.createCell(Column.DON_CHINH.getIndex(), CellType.NUMERIC);
        cell.setCellValue(Column.DON_CHINH.getName());
        cell.setCellStyle(titleStyle);
        // Lấy thêm lần 1
        cell = row.createCell(Column.L1.getIndex(), CellType.NUMERIC);
        cell.setCellValue(Column.L1.getName());
        cell.setCellStyle(titleStyle);
        // Lấy thêm lần 2
        cell = row.createCell(Column.L2.getIndex(), CellType.NUMERIC);
        cell.setCellValue(Column.L2.getName());
        cell.setCellStyle(titleStyle);
        // Tổng
        cell = row.createCell(Column.TONG.getIndex(), CellType.NUMERIC);
        cell.setCellValue(Column.TONG.getName());
        cell.setCellStyle(titleStyle);
        // Đơn giá
        cell = row.createCell(Column.DON_GIA.getIndex(), CellType.STRING);
        cell.setCellValue(Column.DON_GIA.getName());
        cell.setCellStyle(titleStyle);
        // Thành tiền
        cell = row.createCell(Column.THANH_TIEN.getIndex(), CellType.NUMERIC);
        cell.setCellValue(Column.THANH_TIEN.getName());
        cell.setCellStyle(titleStyle);
        // Lợi nhuận
        cell = row.createCell(Column.LOI_NHUAN.getIndex(), CellType.NUMERIC);
        cell.setCellValue(Column.LOI_NHUAN.getName());
        cell.setCellStyle(titleStyle);

        // Data
        List<OrderItem> items = new ArrayList<>();
        for(int i = 0; i < tableModel.getRowCount(); i ++) {
            boolean isEndOfTable = i == tableModel.getRowCount() - 1;
            rownum++;
            row = sheet.createRow(rownum);

            // STT
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue(String.valueOf(tableModel.getValueAt(i, Column.STT.getIndex())));
            cell.setCellStyle(normalCellStyle);
            // Tên sản phẩm
            cell = row.createCell(Column.TEN_SAN_PHAM.getIndex(), CellType.STRING);
            String tenSp = String.valueOf(tableModel.getValueAt(i, Column.TEN_SAN_PHAM.getIndex()));
            cell.setCellValue(tenSp);
            cell.setCellStyle(isEndOfTable ? normalCellStyle : productNameCellStyle);
            // Đơn chính
            cell = row.createCell(Column.DON_CHINH.getIndex(), CellType.NUMERIC);
            Integer quantity = (Integer)tableModel.getValueAt(i, Column.DON_CHINH.getIndex());
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.DON_CHINH.getIndex()));
            cell.setCellStyle(normalCellStyle);
            // Lấy thêm lần 1
            cell = row.createCell(Column.L1.getIndex(), CellType.NUMERIC);
            Integer l1 = (Integer)tableModel.getValueAt(i, Column.L1.getIndex());
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.L1.getIndex()));
            cell.setCellStyle(normalCellStyle);
            // Lấy thêm lần 2
            cell = row.createCell(Column.L2.getIndex(), CellType.NUMERIC);
            Integer l2 = (Integer)tableModel.getValueAt(i, Column.L2.getIndex());
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.L2.getIndex()));
            cell.setCellStyle(normalCellStyle);
            // Tổng
            cell = row.createCell(Column.TONG.getIndex(), CellType.NUMERIC);
            Integer sumQty = (Integer)tableModel.getValueAt(i, Column.TONG.getIndex());
            cell.setCellValue(quantity);
            cell.setCellStyle(isEndOfTable ? sumCellStyle : normalCellStyle);
            // Đơn giá
            cell = row.createCell(Column.DON_GIA.getIndex(), CellType.STRING);
            String unitPrice = String.valueOf(tableModel.getValueAt(i, Column.DON_GIA.getIndex()));
            cell.setCellValue(unitPrice);
            cell.setCellStyle(normalCellStyle);
            // Thành tiền
            cell = row.createCell(Column.THANH_TIEN.getIndex(), CellType.NUMERIC);
            Integer price = Integer.valueOf(tableModel.getValueAt(i, Column.THANH_TIEN.getIndex()).toString());
            cell.setCellValue(price);
            cell.setCellStyle(isEndOfTable ? sumCellStyle : normalCellStyle);
            // Lợi nhuận
            Integer profit = (Integer)tableModel.getValueAt(i, Column.LOI_NHUAN.getIndex());
            cell = row.createCell(Column.LOI_NHUAN.getIndex(), CellType.NUMERIC);
            cell.setCellValue(profit);
            cell.setCellStyle(isEndOfTable ? sumCellStyle : normalCellStyle);
            if(!isEndOfTable) {
                OrderItem item = new OrderItem();
                Product product = new Product();
                product.setName(tenSp);
                item.setProduct(product);
                item.setQuantity(quantity);
                item.setL1(l1);
                item.setL2(l2);
                item.setPrice(price);
                item.setUnitPrice(Integer.valueOf(unitPrice));
                item.setProfit(profit);
                items.add(item);
            }
        }

        Controller.exportOrderToExcel(workbook, fileName + " (lợi nhuận)", true);
        Controller.exportOrderToExcel(workbook, fileName, false);


        //fill order data
        if(Page.ORDER.equals(page)) {
            order.setName(fileName);
            order.setName(fileName);
            Customer customer = new Customer();
            customer.setId(1);
            order.setCustomer(customer);
            order.setCreatedDate(LocalDate.now());
            order.setItems(items);
            processor.insertBillRecord(order);
        }
    }

    public static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)11);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    public static HSSFCellStyle createStyleForNormalCell(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)11);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    public static HSSFCellStyle createStyleForProductName(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)11);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.TAN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    public static void exportOrderToExcel(HSSFWorkbook workbook, String fileName, boolean includedProfit) throws IOException {
        if(!includedProfit) {
            HSSFSheet sheet = workbook.getSheetAt(0);
            Iterator rowIter = sheet.iterator();
            while (rowIter.hasNext()) {
                HSSFRow row = (HSSFRow)rowIter.next();
                HSSFCell cell = row.getCell(Column.LOI_NHUAN.getIndex());
                row.removeCell(cell);
            }
        }
        String filePath = System.getProperty("user.dir");
        if(filePath.contains("/")) {
            filePath += "/";
        } else {
            filePath += "\\";
        }
        File file = new File( filePath + fileName +".xls");
        file.setWritable(true, false);
        file.getParentFile().mkdirs();
        FileOutputStream outFile = new FileOutputStream(file);

        workbook.write(outFile);
        outFile.close();

    }
}
