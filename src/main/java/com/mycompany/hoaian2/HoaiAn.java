/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.hoaian2;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author Jack243
 */
public class HoaiAn extends javax.swing.JFrame {
    /**
     * Creates new form HoaiAn
     */
    public HoaiAn() {
        initComponents();
        setLocationRelativeTo(null);
        setTitle("Chúc vợ yêu dấu làm việc vui vẻ!!!");
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        filePathTbx = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        selectFileBtn = new javax.swing.JButton();
        statusText = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        productList = new javax.swing.JList<>();
        saleLevelCbx = new javax.swing.JComboBox<>();
        jScrollPane2 = new javax.swing.JScrollPane();
        resultTbl = new javax.swing.JTable();
        createTableBtn = new javax.swing.JButton();
        exportBtn = new javax.swing.JButton();
        errorLbl = new javax.swing.JLabel();
        exportFileNameTxt = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        buyLevelCbx = new javax.swing.JComboBox<>();
        profitCbx = new javax.swing.JCheckBox();
        jLabel3 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        jLabel1.setText("File:");

        selectFileBtn.setText("Chọn...");
        selectFileBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                selectFileBtnActionPerformed(evt);
            }
        });

        statusText.setFont(new java.awt.Font("Arial", 1, 12)); // NOI18N
        statusText.setForeground(new java.awt.Color(51, 102, 255));

        productList.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        jScrollPane1.setViewportView(productList);

        saleLevelCbx.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        saleLevelCbx.setMaximumRowCount(20);
        saleLevelCbx.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Mức bán" }));
        saleLevelCbx.setToolTipText("");
        saleLevelCbx.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saleLevelCbxActionPerformed(evt);
            }
        });

        resultTbl.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        resultTbl.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "", "Tên sản phẩm", "Đơn chính", "Lấy thêm L1", "Lấy thêm L2", "Tổng", "Đơn giá", "Thành tiền", "Lợi nhuận"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.String.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, true, true, true, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        resultTbl.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                resultTblPropertyChange(evt);
            }
        });
        jScrollPane2.setViewportView(resultTbl);

        createTableBtn.setText("Tạo bảng");
        createTableBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                createTableBtnActionPerformed(evt);
            }
        });

        exportBtn.setText("Xuất file");
        exportBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportBtnActionPerformed(evt);
            }
        });

        errorLbl.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        errorLbl.setForeground(new java.awt.Color(255, 0, 51));
        errorLbl.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        errorLbl.setToolTipText("");

        jLabel2.setText("Tên file:");

        buyLevelCbx.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        buyLevelCbx.setMaximumRowCount(20);
        buyLevelCbx.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Mức nhập" }));
        buyLevelCbx.setToolTipText("");
        buyLevelCbx.setEnabled(false);
        buyLevelCbx.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                buyLevelCbxActionPerformed(evt);
            }
        });

        profitCbx.setText("In lợi nhuận");

        jLabel3.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel3.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel3.setText("Sản phẩm độc quyền của An Mocha");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(saleLevelCbx, javax.swing.GroupLayout.PREFERRED_SIZE, 119, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(buyLevelCbx, javax.swing.GroupLayout.PREFERRED_SIZE, 124, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addGap(7, 7, 7)
                        .addComponent(jLabel1)
                        .addGap(10, 10, 10)
                        .addComponent(filePathTbx, javax.swing.GroupLayout.PREFERRED_SIZE, 222, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 771, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(createTableBtn)
                            .addComponent(selectFileBtn, javax.swing.GroupLayout.PREFERRED_SIZE, 77, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(statusText, javax.swing.GroupLayout.PREFERRED_SIZE, 146, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 538, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(errorLbl, javax.swing.GroupLayout.PREFERRED_SIZE, 260, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jLabel2)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(exportFileNameTxt)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(profitCbx)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(exportBtn)))))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(filePathTbx, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(selectFileBtn)
                        .addComponent(statusText)
                        .addComponent(jLabel3))
                    .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(saleLevelCbx)
                    .addComponent(buyLevelCbx)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(exportBtn)
                            .addComponent(exportFileNameTxt, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(profitCbx)
                            .addComponent(errorLbl)
                            .addComponent(jLabel2))
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addComponent(createTableBtn, javax.swing.GroupLayout.DEFAULT_SIZE, 27, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 475, Short.MAX_VALUE)
                    .addComponent(jScrollPane1))
                .addContainerGap())
        );

        statusText.getAccessibleContext().setAccessibleName("statusLbl");

        pack();
    }// </editor-fold>//GEN-END:initComponents
    File selectedFile;
    String fileDirectoryPath;
    private void selectFileBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_selectFileBtnActionPerformed
        JFileChooser fc = new JFileChooser();
        int x = fc.showOpenDialog(null);
        if(x == JFileChooser.APPROVE_OPTION) {
            char cbuf[] = null;
            FileReader r;
            selectedFile = fc.getSelectedFile();
            fileDirectoryPath = selectedFile.getAbsolutePath().substring(0, selectedFile.getAbsolutePath().lastIndexOf("/") + 1);
            if(fileDirectoryPath.isEmpty()) {
                fileDirectoryPath = selectedFile.getAbsolutePath().substring(0, selectedFile.getAbsolutePath().lastIndexOf("\\") + 1);
            }
            filePathTbx.setText(selectedFile.getAbsolutePath());
            FileInputStream fis;
            try {
                fis = new FileInputStream(selectedFile);
                readInputExcel(fis);
            } catch (Exception ex) {
                statusText.setText("Please choose again...");
                Logger.getLogger(HoaiAn.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }//GEN-LAST:event_selectFileBtnActionPerformed

    private void createTableBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_createTableBtnActionPerformed
        // TODO add your handling code here:
        errorLbl.setText("");
        fillProductToTable();
    }//GEN-LAST:event_createTableBtnActionPerformed

    

    private void fillProductToTable() {
        productToBuyPrice.clear();
        DefaultTableModel tableModel = (DefaultTableModel)resultTbl.getModel();
        if(!isValidLevel()) {
            return;
        } else {
            int rowCount = tableModel.getRowCount();
            System.out.println("row count:" + rowCount);
            for(int i = rowCount - 1; i >= 0; i--) {
                tableModel.removeRow(i);
            }
        }
        try {
            
            String saleLevel = String.valueOf(saleLevelCbx.getSelectedItem());
            String buyLevel = String.valueOf(buyLevelCbx.getSelectedItem());
            List<String> products = productList.getSelectedValuesList();
            int stt = 0;
            for(String productName : products) {
                stt++;
                tableModel.addRow(new Object[]{stt, productName, 0, 0, 0, 0, productToPrice.get(productName).get(saleLevel), 0, 0});
                productToBuyPrice.put(productName, productToPrice.get(productName).get(buyLevel));
            }
            tableModel.addRow(new Object[]{"", "TỔNG", 0, 0, 0, 0, "", 0, 0});
        } catch (Exception e) {
            errorLbl.setText("Ây da, phần mềm của chồng có lỗi rồi :(");
        }
        
    }
    
    private void resultTblPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_resultTblPropertyChange
        // TODO add your handling code here:
        updateTableValues();
    }//GEN-LAST:event_resultTblPropertyChange

    private void exportBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportBtnActionPerformed
        // TODO add your handling code here:
        errorLbl.setText("");
        try {
           if(exportFileNameTxt.getText().isEmpty()) {
               errorLbl.setText("Vợ yêu quên nhập tên file kìa :)");
           } else {
               exportFile();
           }
        } catch (Exception e) {
            errorLbl.setText("Oạch!!! Bị lỗi gì rồi vợ ơi, vợ làm lại thử xem...");
            Logger.getLogger(HoaiAn.class.getName()).log(Level.SEVERE, null, e);
        }
        
    }//GEN-LAST:event_exportBtnActionPerformed

    private void saleLevelCbxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saleLevelCbxActionPerformed
        // TODO add your handling code here:
        if(isValidLevel() && resultTbl.getRowCount() > 0) {
            updatePriceValueToTable(String.valueOf(saleLevelCbx.getSelectedItem()), true);
        }
    }//GEN-LAST:event_saleLevelCbxActionPerformed

    private void buyLevelCbxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_buyLevelCbxActionPerformed
        // TODO add your handling code here:
        if(isValidLevel() && resultTbl.getRowCount() > 0) {
            updatePriceValueToTable(String.valueOf(buyLevelCbx.getSelectedItem()), false);
        }
    }//GEN-LAST:event_buyLevelCbxActionPerformed
//    static boolean printProfit = false;
    private void updateTableValues() {
        TableModel tableModel = resultTbl.getModel();
        int donchinhTotal = 0;
        int lan1Total = 0;
        int lan2Total = 0;
        int tongTotal = 0;
        int thanhtienTotal = 0;
        int loinhuanTotal = 0;
        for(int row = 0; row < tableModel.getRowCount(); row ++) {
            if(row < (tableModel.getRowCount() - 1)) {
                int donchinh = (Integer)tableModel.getValueAt(row, Column.DON_CHINH.getIndex());
                donchinhTotal += donchinh;
                int lan1 = (Integer)tableModel.getValueAt(row, Column.L1.getIndex());
                lan1Total += lan1;
                int lan2 = (Integer)tableModel.getValueAt(row, Column.L2.getIndex());
                lan2Total += lan2;
                int tong = donchinh + lan1 + lan2;
                tableModel.setValueAt(tong, row, Column.TONG.getIndex());
                tongTotal += tong;
                int dongia = (Integer)tableModel.getValueAt(row, Column.DON_GIA.getIndex());
                int thanhtien = tong * dongia;
                tableModel.setValueAt(thanhtien, row, Column.THANH_TIEN.getIndex());
                thanhtienTotal += thanhtien;
                int giaMua = productToBuyPrice.get(String.valueOf(tableModel.getValueAt(row, Column.TEN_SAN_PHAM.getIndex())));
                int loinhuan = thanhtien - (tong * giaMua);
                tableModel.setValueAt(loinhuan, row, Column.LOI_NHUAN.getIndex());
                loinhuanTotal += loinhuan;
            } else {
                tableModel.setValueAt(donchinhTotal, row, Column.DON_CHINH.getIndex());
                tableModel.setValueAt(lan1Total, row, Column.L1.getIndex());
                tableModel.setValueAt(lan2Total, row, Column.L2.getIndex());
                tableModel.setValueAt(tongTotal, row, Column.TONG.getIndex());
                tableModel.setValueAt(thanhtienTotal, row, Column.THANH_TIEN.getIndex());
                tableModel.setValueAt(loinhuanTotal, row, Column.LOI_NHUAN.getIndex());
            }
        }
    }

    private void updatePriceValueToTable(String level, boolean isSaleLevelChanged) {
        TableModel tableModel = resultTbl.getModel();
        if(!isSaleLevelChanged) {
            productToBuyPrice.clear();
        }
        for(int row = 0; row < tableModel.getRowCount(); row ++) {
            if(row < (tableModel.getRowCount() - 1)) {
                String productName = (String)tableModel.getValueAt(row, Column.TEN_SAN_PHAM.getIndex());
                if(isSaleLevelChanged) {
                    tableModel.setValueAt(productToPrice.get(productName).get(level), row, Column.DON_GIA.getIndex());
                } else {
                    productToBuyPrice.put(productName, productToPrice.get(productName).get(level));
                }

            }
        }
        updateTableValues();
    }

    
    private void readInputExcel(FileInputStream fis) throws IOException {
        while(saleLevelCbx.getItemCount()!= 1) {
            saleLevelCbx.removeItemAt(1);
            buyLevelCbx.removeItemAt(1);
        }
        indexToLevel = new HashMap<>();
        productToPrice = new LinkedHashMap<>();
        HSSFWorkbook wb = new HSSFWorkbook(fis);
        HSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator(); 
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            Iterator<Cell> cellIterator = row.cellIterator();
            String productName = null;
            Map<String, Integer> levelToPrice = new HashMap<>();
            while(cellIterator.hasNext()) {
                
                Cell cell = cellIterator.next();
                CellType cellType = cell.getCellTypeEnum();
                switch (cellType) {
                    case _NONE:
                        System.out.print("");
                        System.out.print("\t");
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        System.out.print("\t");
                        break;
                    case BLANK:
                        System.out.print("");
                        System.out.print("\t");
                        break;
                    case FORMULA:

                        // Công thức
                        System.out.print(cell.getCellFormula());
                        System.out.print("\t");

                        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

                        // In ra giá trị từ công thức
                        System.out.print(evaluator.evaluate(cell).getNumberValue());
                        break;
                    case NUMERIC:
                        System.out.print(Double.valueOf(cell.getNumericCellValue()).intValue());
                        System.out.print("\t");
                        if(row.getRowNum() == 0) {
                            indexToLevel.put(cell.getColumnIndex(), String.valueOf(Double.valueOf(cell.getNumericCellValue()).intValue()));
                        } else {
                            levelToPrice.put(indexToLevel.get(cell.getColumnIndex()), Double.valueOf(cell.getNumericCellValue()).intValue());
                        }
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        System.out.print("\t");
                        if(row.getRowNum() == 0) {
                            indexToLevel.put(cell.getColumnIndex(), cell.getStringCellValue());
                        } else {
                            productName = cell.getStringCellValue();
                        }
                        break;
                    case ERROR:
                        System.out.print("!");
                        System.out.print("\t");
                        break;
               }
            }
            productToPrice.put(productName, levelToPrice);
             System.out.println("");
        }
       initGUI();
    }
    
    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
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
    
    private static HSSFCellStyle createStyleForNormalCell(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)11);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private static HSSFCellStyle createStyleForProductName(HSSFWorkbook workbook) {
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

    private static HSSFCellStyle createStyleForSum(HSSFWorkbook workbook) {
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
    
    private void exportFile() throws FileNotFoundException, IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("don hang");
        sheet.setColumnWidth(Column.STT.getIndex(), 1500);
        sheet.setColumnWidth(Column.TEN_SAN_PHAM.getIndex(), 4500);
        sheet.setColumnWidth(Column.DON_CHINH.getIndex(), 3500);
        sheet.setColumnWidth(Column.L1.getIndex(), 3500);
        sheet.setColumnWidth(Column.L2.getIndex(), 3500);
        sheet.setColumnWidth(Column.TONG.getIndex(), 3500);
        sheet.setColumnWidth(Column.DON_GIA.getIndex(), 3500);
        sheet.setColumnWidth(Column.THANH_TIEN.getIndex(), 3500);
        if(profitCbx.isSelected()) {
            sheet.setColumnWidth(Column.LOI_NHUAN.getIndex(), 3500);
        }

        int rownum = 0;
        Cell cell;
        Row row;
        //
        HSSFCellStyle titleStyle = createStyleForTitle(workbook);
        HSSFCellStyle normalCellStyle = createStyleForNormalCell(workbook);
        HSSFCellStyle productNameCellStyle = createStyleForProductName(workbook);
        HSSFCellStyle sumCellStyle = createStyleForSum(workbook);
 
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
        if(profitCbx.isSelected()) {
            cell = row.createCell(Column.LOI_NHUAN.getIndex(), CellType.NUMERIC);
            cell.setCellValue(Column.LOI_NHUAN.getName());
            cell.setCellStyle(titleStyle);
        }
 
        // Data
        TableModel tableModel = resultTbl.getModel();
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
            cell.setCellValue(String.valueOf(tableModel.getValueAt(i, Column.TEN_SAN_PHAM.getIndex())));
            cell.setCellStyle(isEndOfTable ? normalCellStyle : productNameCellStyle);

            // Đơn chính
            cell = row.createCell(Column.DON_CHINH.getIndex(), CellType.NUMERIC);
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.DON_CHINH.getIndex()));
            cell.setCellStyle(normalCellStyle);
            // Lấy thêm lần 1
            cell = row.createCell(Column.L1.getIndex(), CellType.NUMERIC);
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.L1.getIndex()));
            cell.setCellStyle(normalCellStyle);
            // Lấy thêm lần 2
            cell = row.createCell(Column.L2.getIndex(), CellType.NUMERIC);
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.L2.getIndex()));
            cell.setCellStyle(normalCellStyle);
            // Tổng
            cell = row.createCell(Column.TONG.getIndex(), CellType.NUMERIC);
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.TONG.getIndex()));
            cell.setCellStyle(isEndOfTable ? sumCellStyle : normalCellStyle);
            // Đơn giá
            cell = row.createCell(Column.DON_GIA.getIndex(), CellType.STRING);
            cell.setCellValue(String.valueOf(tableModel.getValueAt(i, Column.DON_GIA.getIndex())));
            cell.setCellStyle(normalCellStyle);
            // Thành tiền
            cell = row.createCell(Column.THANH_TIEN.getIndex(), CellType.NUMERIC);
            cell.setCellValue((Integer)tableModel.getValueAt(i, Column.THANH_TIEN.getIndex()));
            cell.setCellStyle(isEndOfTable ? sumCellStyle : normalCellStyle);
            // Lợi nhuận
            if(profitCbx.isSelected()) {
                cell = row.createCell(Column.LOI_NHUAN.getIndex(), CellType.NUMERIC);
                cell.setCellValue((Integer)tableModel.getValueAt(i, Column.LOI_NHUAN.getIndex()));
                cell.setCellStyle(isEndOfTable ? sumCellStyle : normalCellStyle);
            }
        }
        String fileName = exportFileNameTxt.getText().isEmpty() ? "don hang moi" : exportFileNameTxt.getText();
        File file = new File(fileDirectoryPath + fileName +".xls");
        file.setWritable(true, false);
        file.getParentFile().mkdirs();
 
        FileOutputStream outFile = new FileOutputStream(file);
        workbook.write(outFile);
        outFile.close();
        errorLbl.setText("Vợ yêu tạo xong đơn hàng rồi nè: " + file.getAbsolutePath());
    }
    
    private void initGUI() {
         statusText.setText("Nhập file thành công!");
         System.out.println(indexToLevel.toString());
         productList.setListData(productToPrice.keySet().toArray(new String[productToPrice.size()]));
         for(String level : indexToLevel.values()) {
             saleLevelCbx.addItem(String.valueOf(level));
             buyLevelCbx.addItem(String.valueOf(level));
         }
    }
    
    private boolean isValidLevel() {
        if(saleLevelCbx.getSelectedIndex() == 0) {
           errorLbl.setText("Vợ chọn MỨC BÁN hàng cho SỈ đi kìa <3"); 
           clearBuyLevel();
           return false;
        }
        buyLevelCbx.setEnabled(true);
        if(buyLevelCbx.getSelectedIndex() == 0) {
            errorLbl.setText("Vợ chọn MỨC NHẬP hàng cho MÌNH đi kìa <3");
            return false;
        }
        return true;
    }

    private void clearBuyLevel() {
        buyLevelCbx.setEnabled(false);
        buyLevelCbx.setSelectedIndex(0);
    }
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(HoaiAn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(HoaiAn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(HoaiAn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(HoaiAn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new HoaiAn().setVisible(true);
            }
        });
    }
    
    static Map<String, Integer> productToBuyPrice = new HashMap<>();
    static Map<Integer, String> indexToLevel = new HashMap<>();
    static Map<String, Map<String, Integer>> productToPrice = new HashMap<>();

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> buyLevelCbx;
    private javax.swing.JButton createTableBtn;
    private javax.swing.JLabel errorLbl;
    private javax.swing.JButton exportBtn;
    private javax.swing.JTextField exportFileNameTxt;
    private javax.swing.JTextField filePathTbx;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JList<String> productList;
    private javax.swing.JCheckBox profitCbx;
    private javax.swing.JTable resultTbl;
    private javax.swing.JComboBox<String> saleLevelCbx;
    private javax.swing.JButton selectFileBtn;
    private javax.swing.JLabel statusText;
    // End of variables declaration//GEN-END:variables
}
