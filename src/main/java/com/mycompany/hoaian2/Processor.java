package com.mycompany.hoaian2;

import com.mycompany.models.Customer;
import com.mycompany.models.Order;
import com.mycompany.models.OrderItem;
import com.mycompany.models.Product;
import com.sun.deploy.util.StringUtils;
import org.apache.commons.collections4.CollectionUtils;

import java.sql.*;
import java.sql.Date;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

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
                System.out.println("Load products data successfully!");
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
            stm.setTimestamp(3, Timestamp.valueOf(order.getCreatedDate().atStartOfDay()));
            stm.executeUpdate();

            sql = "INSERT INTO orderitems(order_id,product_id,quantity,l1,l2,price,unitprice,profit) " +
                    "SELECT orders.id, ?, ?, ?, ?, ?, ?, ?" +
                    "  FROM orders" +
                    " WHERE orders.name = ?" +
                    " LIMIT 1";
            stm = conn.prepareStatement(sql);
            for (OrderItem item : order.getItems()) {
                stm.setInt(1, item.getProduct().getId());
                stm.setInt(2, item.getQuantity());
                stm.setInt(3, item.getL1());
                stm.setInt(4, item.getL2());
                stm.setInt(5, item.getPrice());
                stm.setInt(6, item.getUnitPrice());
                stm.setInt(7, item.getProfit());
                stm.setString(8, order.getName());
                stm.addBatch();
            }
            stm.executeBatch();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            dataConnector.cleanUp(conn, stm, rs);
        }
    }

    public List<Order> getOrdersByDateRange(String fromDate, String toDate) {
        Connection conn = null;
        Statement stm = null;
        ResultSet rs = null;
        DateTimeFormatter formatter = DateTimeFormatter.ISO_DATE;
        String sDate = LocalDate.parse(fromDate).atStartOfDay().atZone(ZoneOffset.UTC).format(DateTimeFormatter.ISO_DATE);
        String eDate = LocalDate.parse(toDate).atTime(LocalTime.MAX).atZone(ZoneOffset.UTC).format(DateTimeFormatter.ISO_DATE);
        Map<Integer, Order> idToOrder = new HashMap<>();
        try {
            String sql = "SELECT orders.id as orderId, orders.name as orderName, created_date, " +
                    "  orderitems.*, " +
                    "  products.name as productName, products.id as productId, " +
                    "  customers.id as customerId, customers.name as customerName, customers.phone, customers.address " +
                    "FROM orders " +
                    "  LEFT JOIN orderitems on orderitems.order_id = orders.id " +
                    "  LEFT JOIN products on products.id = orderitems.product_id " +
                    "  LEFT JOIN customers on orders.customer_id = customers.id " +
                    "where created_date >= '" + sDate + " 00:00:00' and created_date <= '" + eDate + " 23:59:59'";
            conn = dataConnector.getMySQLConnection();
            stm = conn.createStatement();
            rs = stm.executeQuery(sql);
            while (rs.next()) {
                Order order = idToOrder.get(rs.getInt("orderId"));
                if(Objects.isNull(order)) {
                    order = new Order();
                    order.setId(rs.getInt("orderId"));
                    order.setName(rs.getString("orderName"));
                    order.setCreatedDate(rs.getTimestamp("created_date").toLocalDateTime().toLocalDate());
                    Customer customer = new Customer();
                    customer.setId(rs.getInt("customerId"));
                    customer.setName(rs.getString("customerName"));
                    customer.setPhone(rs.getString("phone"));
                    customer.setAddress(rs.getString("address"));
                    order.setCustomer(customer);
                    order.setItems(new ArrayList<>());
                }
                OrderItem item = new OrderItem();
                item.setOrderId(order.getId());
                item.setUnitPrice(rs.getInt("unitprice"));
                item.setQuantity(rs.getInt("quantity"));
                item.setL1(rs.getInt("l1"));
                item.setL2(rs.getInt("l2"));
                item.setPrice(rs.getInt("price"));
                item.setProfit(rs.getInt("profit"));
                Product product = new Product();
                product.setId(rs.getInt("product_id"));
                item.setProduct(product);
                order.getItems().add(item);
                idToOrder.put(order.getId(), order);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return new ArrayList<>(idToOrder.values());
    }

    public void deleteOrders(List<Order> orders) {
        Connection conn = null;
        Statement stm = null;
        ResultSet rs = null;
        List<String> orderIdsToDelete = orders.stream().map(o -> o.getId().toString()).collect(Collectors.toList());
        try {
            String sql = "DELETE FROM orders WHERE id IN (" + StringUtils.join(orderIdsToDelete,",") + ")";
            conn = dataConnector.getMySQLConnection();
            stm = conn.createStatement();
            stm.executeUpdate(sql);

            sql = "DELETE FROM orderitems WHERE order_id IN (" + StringUtils.join(orderIdsToDelete,",") + ")";
            stm.executeUpdate(sql);
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            dataConnector.cleanUp(conn, stm, rs);
        }
    }

    private static java.sql.Date getCurrentDate() {
        java.util.Date today = new java.util.Date();
        return new java.sql.Date(today.getTime());
    }

    public Order summarizeOrders(List<Order> orders) {
        if(CollectionUtils.isEmpty(orders)) {
            return null;
        }
        Order finalOrder = null;
        try {
            finalOrder = orders.get(0).clone();
            Map<Integer, OrderItem> productIdToItem = new HashMap<>();
            for(Order order : orders) {
                for(OrderItem item : order.getItems()) {
                    OrderItem existedItem = productIdToItem.get(item.getProduct().getId());
                    if(Objects.isNull(existedItem)) {
                        existedItem = item.clone();
                    } else {
                        existedItem.setQuantity(existedItem.getQuantity() + item.getQuantity());
                        existedItem.setL1(existedItem.getL1() + item.getL1());
                        existedItem.setL2(existedItem.getL2() + item.getL2());
                        existedItem.setPrice(existedItem.getPrice() + item.getPrice());
                        existedItem.setProfit(existedItem.getProfit() + item.getProfit());
                    }
                    productIdToItem.put(existedItem.getProduct().getId(), existedItem);
                }
            }
            finalOrder.setItems(new ArrayList<>(productIdToItem.values()));
        } catch (CloneNotSupportedException e) {
            e.printStackTrace();
        } finally {
            return finalOrder;
        }
    }
}
