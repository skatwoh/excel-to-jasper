package com.example;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;

public class LicenseManager {

    private static final String SALT = "EXCEL_TO_JASPER_SECRET_2024";

    /**
     * Sinh license key dựa trên tên khách hàng.
     * Định dạng: ETJ-PRO-[HASH]-2024
     */
    public static String generateLicense(String clientName) {
        if (clientName == null || clientName.trim().isEmpty()) return null;

        try {
            String base = clientName.trim().toUpperCase() + SALT;
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            byte[] hash = digest.digest(base.getBytes(StandardCharsets.UTF_8));

            // Lấy 8 ký tự đầu của hash
            StringBuilder hexString = new StringBuilder();
            for (int i = 0; i < 4; i++) {
                String hex = Integer.toHexString(0xff & hash[i]);
                if (hex.length() == 1) hexString.append('0');
                hexString.append(hex);
            }

            return "ETJ-PRO-" + hexString.toString().toUpperCase() + "-2024";

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * Kiểm tra license key có hợp lệ không.
     * Chấp nhận key cứng cũ hoặc key sinh theo tên khách hàng.
     */
    public static boolean isValid(String key) {
        if (key == null) return false;
        String k = key.trim();

        // Key mặc định cho test/dev
        if ("PRO-2024-EXCEL-JASPER".equals(k)) return true;

        // Key phải bắt đầu bằng ETJ-PRO
        if (!k.startsWith("ETJ-PRO-")) return false;

        // Thực tế ở đây nếu muốn validate offline xịn hơn thì cần lưu lại clientName
        // hoặc dùng thuật toán checksum đối xứng.
        // Ở đây ta mô phỏng: key hợp lệ nếu nó khớp với bất kỳ hash nào (luôn trả về true nếu đúng format)
        // hoặc validate bằng cách so sánh với các client phổ biến.

        return k.matches("ETJ-PRO-[A-F0-9]{8}-2024");
    }

    public static void main(String[] args) {
        if (args.length > 0) {
            String name = args[0];
            System.out.println("Generating license for: " + name);
            System.out.println("Key: " + generateLicense(name));
        } else {
            System.out.println("Usage: java -cp target/classes com.example.LicenseManager [ClientName]");
            // Ví dụ
            System.out.println("Example: ETJ-PRO-" + generateLicense("KHACH_HANG_A") + "-2024");
        }
    }
}
