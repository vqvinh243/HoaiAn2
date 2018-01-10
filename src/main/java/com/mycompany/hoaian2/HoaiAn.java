/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.hoaian2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

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
        levelCbx = new javax.swing.JComboBox<>();
        jScrollPane2 = new javax.swing.JScrollPane();
        resultTbl = new javax.swing.JTable();
        createTableBtn = new javax.swing.JButton();
        exportBtn = new javax.swing.JButton();
        errorLbl = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        jLabel1.setText("File:");

        selectFileBtn.setText("Chọn...");
        selectFileBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                selectFileBtnActionPerformed(evt);
            }
        });

        productList.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        jScrollPane1.setViewportView(productList);

        levelCbx.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        levelCbx.setMaximumRowCount(20);
        levelCbx.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Mức nhập" }));
        levelCbx.setToolTipText("");

        resultTbl.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        resultTbl.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "", "Tên sản phẩm", "Đơn chính", "Lấy thêm L1", "Lấy thêm L2", "Tổng", "Đơn giá", "Thành tiền"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.String.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class, java.lang.Integer.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, true, true, true, false, false, false
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

        errorLbl.setFont(new java.awt.Font("Arial", 0, 11)); // NOI18N
        errorLbl.setForeground(new java.awt.Color(255, 0, 51));

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(7, 7, 7)
                        .addComponent(jLabel1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(filePathTbx, javax.swing.GroupLayout.PREFERRED_SIZE, 272, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(selectFileBtn)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(statusText))
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(levelCbx, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(createTableBtn))
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.PREFERRED_SIZE, 158, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(errorLbl, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addGap(18, 18, 18)
                                .addComponent(exportBtn))
                            .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 645, Short.MAX_VALUE))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(filePathTbx, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(selectFileBtn)
                        .addComponent(statusText))
                    .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(levelCbx, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(createTableBtn)
                    .addComponent(exportBtn)
                    .addComponent(errorLbl))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jScrollPane2)
                    .addComponent(jScrollPane1))
                .addContainerGap())
        );

        statusText.getAccessibleContext().setAccessibleName("statusLbl");

        pack();
    }// </editor-fold>//GEN-END:initComponents
    File selectedFile;
    private void selectFileBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_selectFileBtnActionPerformed
        JFileChooser fc = new JFileChooser();
        int x = fc.showOpenDialog(null);
        if(x == JFileChooser.APPROVE_OPTION) {
            char cbuf[] = null;
            FileReader r;
            selectedFile = fc.getSelectedFile();
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
        if(!isValidLevel()) {
            errorLbl.setText("Hãy chọn mức nhập!!!");
            return;
        }
        try {
            Integer level = Integer.parseInt((String)levelCbx.getSelectedItem());
            List<String> products = productList.getSelectedValuesList();
            DefaultTableModel tableModel = (DefaultTableModel)resultTbl.getModel();
            int stt = 0;
            for(String productName : products) {
                stt++;
                tableModel.addRow(new Object[]{stt, productName, 0, 0, 0, 0, productToPrice.get(productName).get(level), 0});
            }
            tableModel.addRow(new Object[]{"", "Tổng", 0, 0, 0, 0, "", 0});
        } catch (Exception e) {
            errorLbl.setText("Đã có lỗi gì đó xảy ra!!!");
        }
        
    }
    
    private void resultTblPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_resultTblPropertyChange
        // TODO add your handling code here:
        updateTableValues();
    }//GEN-LAST:event_resultTblPropertyChange

    private void updateTableValues() {
        TableModel tableModel = resultTbl.getModel();
        int donchinhTotal = 0;
        int lan1Total = 0;
        int lan2Total = 0;
        int tongTotal = 0;
        int thanhtienTotal = 0;
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
            } else {
                tableModel.setValueAt(donchinhTotal, row, Column.DON_CHINH.getIndex());
                tableModel.setValueAt(lan1Total, row, Column.L1.getIndex());
                tableModel.setValueAt(lan2Total, row, Column.L2.getIndex());
                tableModel.setValueAt(tongTotal, row, Column.TONG.getIndex());
                tableModel.setValueAt(thanhtienTotal, row, Column.THANH_TIEN.getIndex());
            }
        }
    }
    
    static Map<Integer, Integer> indexToLevel = new HashMap<>();
    static Map<String, Map<Integer, Integer>> productToPrice = new HashMap<>();
    private void readInputExcel(FileInputStream fis) throws IOException {
        indexToLevel = new HashMap<>();
        productToPrice = new HashMap<>();
        HSSFWorkbook wb = new HSSFWorkbook(fis);
        HSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator(); 
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            Iterator<Cell> cellIterator = row.cellIterator();
            String productName = null;
            Map<Integer, Integer> levelToPrice = new HashMap<>();
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
                            indexToLevel.put(cell.getColumnIndex(), Double.valueOf(cell.getNumericCellValue()).intValue());
                        } else {
                            levelToPrice.put(indexToLevel.get(cell.getColumnIndex()), Double.valueOf(cell.getNumericCellValue()).intValue());
                        }
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        System.out.print("\t");
                        if(row.getRowNum() > 0) {
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
    
    private void initGUI() {
         statusText.setText("Done!");
         System.out.println(productToPrice.toString());
         productList.setListData(productToPrice.keySet().toArray(new String[productToPrice.size()]));
         for(Integer level : indexToLevel.keySet()) {
             levelCbx.addItem(String.valueOf(level));
         }
    }
    
    private boolean isValidLevel() {
        try {
            Integer.parseInt((String)levelCbx.getSelectedItem());
        } catch (NumberFormatException e) {
            return false;
        }
        return true;
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

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton createTableBtn;
    private javax.swing.JLabel errorLbl;
    private javax.swing.JButton exportBtn;
    private javax.swing.JTextField filePathTbx;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JComboBox<String> levelCbx;
    private javax.swing.JList<String> productList;
    private javax.swing.JTable resultTbl;
    private javax.swing.JButton selectFileBtn;
    private javax.swing.JLabel statusText;
    // End of variables declaration//GEN-END:variables
}
