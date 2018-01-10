/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.hoaian2;

/**
 *
 * @author Jack243
 */
public enum Column {
    STT("",0),
    TEN_SAN_PHAM("Tên sản phẩm",1),
    DON_CHINH("Đơn chính",2),
    L1("Lấy thêm L1",3),
    L2("Lấy thêm L2",4),
    TONG("Tổng",5),
    DON_GIA("Đơn giá",6),
    THANH_TIEN("Thành tiền",7);
    
    String name;
    Integer index;
    
    private Column(String name, Integer index) {
        this.name = name;
        this.index = index;
    }
    
    public String getName() {
        return this.name;
    }
    
    public Integer getIndex() {
        return this.index;
    }
}
