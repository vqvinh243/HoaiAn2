package com.mycompany.hoaian2;

import com.zaxxer.hikari.HikariDataSource;

import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DataConnector {
    private HikariDataSource mySQLDS;

    private static class QueryUtilsHelper {
        private static final DataConnector INSTANCE = new DataConnector();
    }

    private DataConnector() {
        try {
            mySQLDS = setMySQLConnection();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    public static DataConnector getInstance(){
        return QueryUtilsHelper.INSTANCE;
    }

    private HikariDataSource setMySQLConnection() throws SQLException {
        HikariDataSource ds = new HikariDataSource();
        ds.setDriverClassName("com.mysql.jdbc.Driver");
        ds.setJdbcUrl("jdbc:mysql://localhost:3306/hoaian");
        ds.setUsername("root");
        ds.setPassword("quangvinh");
        ds.setMaximumPoolSize(5);
        ds.setLeakDetectionThreshold(1000000);
        return ds;
    }

    public List<Map<String, Object>> convertResultSetToList(final ResultSet rs) throws SQLException {
        final ResultSetMetaData md = rs.getMetaData();
        final int columns = md.getColumnCount();
        final List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();

        while (rs.next()) {
            final Map<String, Object> row = new HashMap<String, Object>(columns);
            for (int i = 1; i <= columns; ++i) {
                row.put(md.getColumnLabel(i), rs.getObject(i));
            }
            list.add(row);
        }

        return list;
    }

    public Connection getMySQLConnection() {
        try {
            if(mySQLDS == null) {
                mySQLDS = setMySQLConnection();
            }
            return mySQLDS.getConnection();
        } catch (SQLException e) {
            e.printStackTrace();
            return null;
        }
    }

    public void cleanUp(Connection conn, Statement stm, ResultSet rs) {
        try {
            if (rs != null) {
                rs.close();
            }
            if (stm != null) {
                stm.close();
            }
            if (conn != null) {
                conn.close();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}
