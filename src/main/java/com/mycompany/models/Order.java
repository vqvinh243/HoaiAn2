package com.mycompany.models;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class Order implements Cloneable{
    Integer id;
    String name;
    Customer customer;
    LocalDate createdDate;
    List<OrderItem> items;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LocalDate getCreatedDate() {
        return createdDate;
    }

    public void setCreatedDate(LocalDate createdDate) {
        this.createdDate = createdDate;
    }

    public List<OrderItem> getItems() {
        return items;
    }

    public void setItems(List<OrderItem> items) {
        this.items = items;
    }

    public Customer getCustomer() {
        return customer;
    }

    public void setCustomer(Customer customer) {
        this.customer = customer;
    }

    @Override
    public Order clone() throws CloneNotSupportedException {
        Order cloned = (Order) super.clone();
        cloned.setItems(new ArrayList<>());
        return cloned;
    }
}
