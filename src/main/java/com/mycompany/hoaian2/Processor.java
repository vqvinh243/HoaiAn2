package com.mycompany.hoaian2;

import com.mycompany.models.Order;
import com.mycompany.models.OrderItem;

import java.sql.*;
import java.util.HashMap;
import java.util.Map;

public class Processor {
    private DataConnector dataConnector = DataConnector.getInstance();
    public static Map<String, Integer> productNameToId = new HashMap<>();
    public Processor() {
        if(productNameToId.isEmpty()) {
            Connection conn = null;
            Statement stm = null;
            ResultSet rs = null;
            try {
                String sql = "SELECT * FROM products";
                conn = dataConnector.getMySQLConnection();
                stm = conn.createStatement();
                rs = stm.executeQuery(sql);
                while (rs.next()) {
                    productNameToId.put(rs.getString("name"), rs.getInt("id"));
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                dataConnector.cleanUp(conn, stm, rs);
            }
        }
    }
    public void insertBillRecord(Order order) {
        Connection conn = null;
        PreparedStatement stm = null;
        ResultSet rs = null;
        try {
            String sql = "INSERT INTO orders(name,customer_id,created_date) " +
                            "VALUES (?, ?, ?)";
            conn = dataConnector.getMySQLConnection();
            stm = conn.prepareStatement(sql);
            stm.setString(1, order.getName());
            stm.setInt(2, order.getCustomer().getId());
            stm.setDate(3, Date.valueOf(order.getCreatedDate()));
            stm.executeUpdate();

            sql = "INSERT INTO orderitems(order_id,product_id,quantity,price,unitprice,profit) " +
                    "SELECT orders.id, ?, ?, ?, ?, ?" +
                    "  FROM orders" +
                    " WHERE orders.name = ?" +
                    " LIMIT 1";
            stm = conn.prepareStatement(sql);
            for (OrderItem item : order.getItems()) {
                stm.setInt(1, item.getProduct().getId());
                stm.setInt(2, item.getQuantity());
                stm.setInt(3, item.getPrice());
                stm.setInt(4, item.getUnitPrice());
                stm.setInt(5, item.getProfit());
                stm.setString(6, order.getName());
                stm.addBatch();
            }
            stm.executeBatch();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            dataConnector.cleanUp(conn, stm, rs);
        }
    }
    private static java.sql.Date getCurrentDate() {
        java.util.Date today = new java.util.Date();
        return new java.sql.Date(today.getTime());
    }
}
