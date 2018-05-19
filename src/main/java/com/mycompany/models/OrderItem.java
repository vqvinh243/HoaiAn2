package com.mycompany.models;

public class OrderItem implements Cloneable{
    Integer id;
    Integer orderId;
    Product product;
    Integer quantity;
    Integer l1;
    Integer l2;
    Integer price;
    Integer unitPrice;
    Integer profit;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public Integer getOrderId() {
        return orderId;
    }

    public void setOrderId(Integer orderId) {
        this.orderId = orderId;
    }

    public Integer getQuantity() {
        return quantity;
    }

    public void setQuantity(Integer quantity) {
        this.quantity = quantity;
    }

    public Integer getL1() {
        return l1;
    }

    public void setL1(Integer l1) {
        this.l1 = l1;
    }

    public Integer getL2() {
        return l2;
    }

    public void setL2(Integer l2) {
        this.l2 = l2;
    }

    public Product getProduct() {
        return product;
    }

    public void setProduct(Product product) {
        this.product = product;
    }

    public Integer getPrice() {
        return price;
    }

    public void setPrice(Integer price) {
        this.price = price;
    }

    public Integer getUnitPrice() {
        return unitPrice;
    }

    public void setUnitPrice(Integer unitPrice) {
        this.unitPrice = unitPrice;
    }

    public Integer getProfit() {
        return profit;
    }

    public void setProfit(Integer profit) {
        this.profit = profit;
    }

    @Override
    public OrderItem clone() throws CloneNotSupportedException {
        OrderItem item = (OrderItem) super.clone();
        return item;
    }
}
