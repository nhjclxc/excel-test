package com.example.exceltest;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class Main {
    public static void main(String[] args) {
        // 假设我们有一个Iterator对象
        Iterator<String> iterator = Arrays.asList("a", "b", "c", "d").iterator();
        // 将Iterator转换为Stream
        List<String> list = StreamSupport.stream( Spliterators.spliteratorUnknownSize(iterator, Spliterator.ORDERED),false).collect(Collectors.toList());
        // 输出结果
        System.out.println(list); // 输出：[a, b, c, d]
    }
}
