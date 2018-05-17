package com.mycompany.models;

import com.mycompany.hoaian2.Processor;

public class Product {
    Integer id;
    String name;

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
        setId(Processor.productNameToId.get(name));
    }
}
