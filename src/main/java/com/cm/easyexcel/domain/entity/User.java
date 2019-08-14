package com.cm.easyexcel.domain.entity;

import com.alibaba.excel.metadata.BaseRowModel;

/**
 * @description:
 * @author: caomian
 * @data: 2019/8/14 9:19
 */
public class User extends BaseRowModel {
    private String name;
    private int age;
    private String gender;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }
}
