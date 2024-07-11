//package com.example.exceltest;
//
//import java.text.ParseException;
//import java.text.SimpleDateFormat;
//import java.time.LocalDate;
//import java.time.LocalDateTime;
//import java.time.LocalTime;
//import java.time.format.DateTimeFormatter;
//import java.util.Date;
//
///**
// * 时间传参全局处理
// */
//public class DateTimeFormatterUtils {
//
//    public static class LocalDateTimeFormatter {
//        private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
//        public static LocalDateTime parse(String text, java.util.Locale locale) {
//            return LocalDateTime.parse(text, dateTimeFormatter);
//        }
//        public static String print(LocalDateTime object, java.util.Locale locale) {
//            return dateTimeFormatter.format(object);
//        }
//    }
//    public static class LocalDateFormatter {
//        private static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//        public static LocalDate parse(String text, java.util.Locale locale) {
//            return LocalDate.parse(text, dateFormatter);
//        }
//        public static String print(LocalDate object, java.util.Locale locale) {
//            return dateFormatter.format(object);
//        }
//    }
//    public static class LocalTimeFormatter {
//        private static final DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
//        public static LocalTime parse(String text, java.util.Locale locale) {
//            return LocalTime.parse(text, timeFormatter);
//        }
//        public static String print(LocalTime object, java.util.Locale locale) {
//            return timeFormatter.format(object);
//        }
//    }
//
//    public static class DateFormatter {
//        private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
//        public static Date parse(String text, java.util.Locale locale) throws ParseException {
//            return dateFormat.parse(text);
//        }
//        public static String print(Date date, java.util.Locale locale) {
//            return dateFormat.format(date);
//        }
//    }
//    public static class DateTimeFormatterCustom {
//        private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
//        public static Date parse(String text, java.util.Locale locale){
//            try {
//                return dateFormat.parse(text);
//            } catch (ParseException ignored) {}
//            return null;
//        }
//        public static String print(Date date, java.util.Locale locale) {
//            return dateFormat.format(date);
//        }
//    }
//}
