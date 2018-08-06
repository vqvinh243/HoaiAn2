package com.mycompany.hoaian2;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.nio.file.Paths;
import java.util.Properties;

public class InstanceProperties {

    private static final Logger logger = LoggerFactory.getLogger(InstanceProperties.class);
    private static final InstanceProperties instanceProps = new InstanceProperties();
    private static Properties props;

    public static boolean isPrivate = false;

    private String title;
    private String selectBuyLevelReminder;
    private String selectSellLevelReminder;
    private String connectInternetReminder;
    private String fileNameInputReminder;
    private String congratulationReminder;
    private String unknownErrorWarning;
    private String deleteConfirmation;

    private InstanceProperties() {
        props = new Properties();
        loadProps();
    }

    public void loadProps() {
        props = new Properties();
        try {
//            logger.info("Property file is getting from {}", Paths.get(".").toAbsolutePath().normalize().toString() + "/instance.props");
//            props.load(new FileInputStream(
//                    Paths.get(".").toAbsolutePath().normalize().toString() + "/instance.props"));
            this.title = isPrivate ? "Chúc vợ yêu dấu làm việc vui vẻ!!!" : "Phần mềm tính đơn hàng Mocha";
            this.selectBuyLevelReminder = isPrivate ? "Vợ chọn MỨC NHẬP hàng cho MÌNH đi kìa <3" : "Hãy chọn MỨC NHẬP HÀNG";
            this.selectSellLevelReminder = isPrivate ? "Vợ chọn MỨC BÁN hàng cho SỈ đi kìa <3" : "Hãy chọn MỨC BÁN HÀNG";
            this.connectInternetReminder = isPrivate ? "Vợ nhớ kết nối mạng khi sử dụng tính năng này nha!" : "Nhớ kết nối internet trước khi dùng tính năng này!";
            this.fileNameInputReminder = isPrivate ? "Vợ yêu quên nhập tên file kìa :)" : "Hãy nhập tên file!!!";
            this.congratulationReminder = isPrivate ? "Vợ yêu tạo xong đơn hàng rồi nè!" : "Đã tạo xong đơn hàng!";
            this.unknownErrorWarning = isPrivate ? "Oạch!!! Bị lỗi gì rồi vợ ơi, vợ làm lại thử xem..." : "Có lỗi xảy ra, xin hãy làm lại";
            this.deleteConfirmation = isPrivate ? "Vợ yêu có thực sự muốn xóa không?" : "Xác nhận xóa đơn hàng!";
        } catch (Exception e) {
            logger.error("Load properties error", e);
        }
    }

    public static InstanceProperties getInstance() {
        return instanceProps;
    }

    public String getTitle() {
        return title;
    }

    public String getSelectBuyLevelReminder() {
        return selectBuyLevelReminder;
    }

    public String getSelectSellLevelReminder() {
        return selectSellLevelReminder;
    }

    public String getConnectInternetReminder() {
        return connectInternetReminder;
    }

    public String getFileNameInputReminder() {
        return fileNameInputReminder;
    }

    public String getCongratulationReminder() {
        return congratulationReminder;
    }

    public String getUnknownErrorWarning() {
        return unknownErrorWarning;
    }

    public String getDeleteConfirmation() {
        return deleteConfirmation;
    }

    public static boolean isPrivate() {
        return isPrivate;
    }
}